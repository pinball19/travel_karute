/**
 * スプレッドシート関連の機能を管理するモジュール
 */
const SpreadsheetManager = {
  // Handsontable インスタンス
  hot: null,
  
  /**
   * Handsontableの初期化
   */
  initialize: function() {
    const container = document.getElementById('hot-container');
    
    // コンテナが見つからない場合は初期化しない
    if (!container) {
      console.error('hot-containerが見つかりません');
      return;
    }
    
    // 既存のインスタンスがあれば破棄
    if (this.hot) {
      this.hot.destroy();
      this.hot = null;
    }
    
    try {
      // 基本設定
      const options = {
        data: UNIFIED_TEMPLATE,
        rowHeaders: true,
        colHeaders: true, // 列ヘッダー（A,B,C...）を表示
        contextMenu: true,
        manualColumnResize: true,
        manualRowResize: true,
        licenseKey: 'non-commercial-and-evaluation',
        mergeCells: UNIFIED_MERGED_CELLS,
        
        // カラム幅を設定
        colWidths: COLUMN_WIDTHS,
        
        // 表示設定
        width: container.offsetWidth,
        height: container.offsetHeight,
        
        // テキスト折り返し
        wordWrap: true,
        
        // 最小表示行数
        minRows: 30,
        minSpareRows: 5,
        
        // 挿入・削除時のフック
        afterCreateRow: function(index, amount) {
          console.log(`${amount}行が位置${index}に挿入されました`);
          SpreadsheetManager.updateAfterRowOperation();
        },
        
        afterRemoveRow: function(index, amount) {
          console.log(`${amount}行が位置${index}から削除されました`);
          SpreadsheetManager.updateAfterRowOperation();
        },
        
        // セルのカスタマイズ
        cells: function(row, col) {
          const cellProperties = {};
          
          // セクション見出し行のスタイル
          if (SECTION_STYLES.sectionHeaders.includes(row)) {
            cellProperties.renderer = SpreadsheetManager.sectionHeaderRenderer;
            cellProperties.readOnly = true; // セクション見出しは編集不可に
          }
          // ヘッダー行のスタイル
          else if (SECTION_STYLES.headerRows.includes(row)) {
            cellProperties.renderer = SpreadsheetManager.headerRenderer;
          }
          // 合計行のスタイル
          else if (SECTION_STYLES.totalRows.includes(row)) {
            cellProperties.renderer = SpreadsheetManager.totalRenderer;
          }
          // 空白の区切り行
          else if (SECTION_STYLES.spacerRows.includes(row)) {
            cellProperties.renderer = SpreadsheetManager.spacerRenderer;
          }
          
          // 1列目（項目名列）の設定
          if (col === 0) {
            cellProperties.wordWrap = true;
          }
          
          return cellProperties;
        },
        
        // 全セルを選択させない
        disableVisualSelection: false,
        
        // セルの編集可否を設定
        readOnly: false,
        
        // セル結合の調整
        afterGetColHeader: function(col, TH) {
          // 列ヘッダーのスタイル調整
        },
        
        // 行の高さ自動調整
        autoRowSize: {
          syncLimit: 1000
        }
      };
      
      console.log('Handsontableを初期化します');
      // Handsontableインスタンスを作成
      this.hot = new Handsontable(container, options);
      console.log('Handsontableの初期化完了');
      
      // 右クリックメニューのカスタマイズ
      this.customizeContextMenu();
      
      // レンダリングを確実に行う
      setTimeout(() => {
        if (this.hot) {
          console.log('Handsontableを再レンダリングします');
          this.hot.render();
          container.scrollTop = 0;
        }
      }, 200);
    } catch (error) {
      console.error('Handsontableの初期化中にエラーが発生しました:', error);
    }
    
    // ウィンドウのリサイズに合わせてHandsontableをリサイズ
    window.addEventListener('resize', () => {
      if (this.hot) {
        console.log('リサイズによりHandsontableを更新します');
        this.hot.updateSettings({
          width: container.offsetWidth,
          height: container.offsetHeight
        });
        this.hot.render();
      }
    });
  },
  
  /**
   * 右クリックメニューのカスタマイズ
   */
  customizeContextMenu: function() {
    if (!this.hot) return;
    
    // 既存のコンテキストメニュー設定を取得
    const contextMenuItems = this.hot.getSettings().contextMenu;
    
    // 行挿入操作の後に背景色を修正するためのラッパー関数
    const customInsertRow = {
      name: '行を挿入',
      callback: function(key, selection) {
        // 既存の行挿入処理を実行
        Handsontable.plugins.ContextMenu.DEFAULT_ITEMS.row_above.callback.call(this, key, selection);
        
        // セクションスタイルの再適用
        SpreadsheetManager.updateAfterRowOperation();
      }
    };
    
    // カスタマイズしたコンテキストメニュー
    const customContextMenu = {
      items: {
        'row_above': customInsertRow,
        'row_below': {
          name: '行を下に挿入',
          callback: function(key, selection) {
            Handsontable.plugins.ContextMenu.DEFAULT_ITEMS.row_below.callback.call(this, key, selection);
            SpreadsheetManager.updateAfterRowOperation();
          }
        },
        'remove_row': {
          name: '行を削除',
          callback: function(key, selection) {
            Handsontable.plugins.ContextMenu.DEFAULT_ITEMS.remove_row.callback.call(this, key, selection);
            SpreadsheetManager.updateAfterRowOperation();
          }
        },
        'separator1': Handsontable.plugins.ContextMenu.DEFAULT_ITEMS.separator1,
        'copy': Handsontable.plugins.ContextMenu.DEFAULT_ITEMS.copy,
        'cut': Handsontable.plugins.ContextMenu.DEFAULT_ITEMS.cut,
        'separator2': Handsontable.plugins.ContextMenu.DEFAULT_ITEMS.separator2,
        'alignment': Handsontable.plugins.ContextMenu.DEFAULT_ITEMS.alignment
      }
    };
    
    // コンテキストメニューを更新
    this.hot.updateSettings({
      contextMenu: customContextMenu
    });
  },
  
  /**
   * 行の操作（挿入/削除）後に背景色などを更新
   */
  updateAfterRowOperation: function() {
    if (!this.hot) return;
    
    // データを取得
    const data = this.hot.getData();
    
    // 既存のマージセル設定を取得
    const currentMerges = this.hot.getSettings().mergeCells || [];
    
    // セクションヘッダーの位置を探し、スタイル設定を更新
    let updatedSectionHeaders = [];
    let updatedHeaderRows = [];
    let updatedTotalRows = [];
    
    // 基本情報セクション
    const basicSectionIndex = 0;
    updatedSectionHeaders.push(basicSectionIndex);
    updatedHeaderRows.push(basicSectionIndex + 1);
    updatedHeaderRows.push(basicSectionIndex + 2);
    
    // 入金情報セクション
    let paymentSectionIndex = -1;
    for (let i = basicSectionIndex + 3; i < data.length; i++) {
      if (data[i][0] === '◆ 入金情報') {
        paymentSectionIndex = i;
        break;
      }
    }
    
    if (paymentSectionIndex > 0) {
      updatedSectionHeaders.push(paymentSectionIndex);
      updatedHeaderRows.push(paymentSectionIndex + 1);
      
      // 入金合計行を探す
      for (let i = paymentSectionIndex + 2; i < data.length; i++) {
        if (data[i][0] === 'A；入金合計') {
          updatedTotalRows.push(i);
          break;
        }
      }
    }
    
    // 支払情報セクション
    let expenseSectionIndex = -1;
    for (let i = paymentSectionIndex + 2; i < data.length; i++) {
      if (data[i][0] === '◆ 支払情報') {
        expenseSectionIndex = i;
        break;
      }
    }
    
    if (expenseSectionIndex > 0) {
      updatedSectionHeaders.push(expenseSectionIndex);
      updatedHeaderRows.push(expenseSectionIndex + 1);
      
      // 支払合計行を探す
      for (let i = expenseSectionIndex + 2; i < data.length; i++) {
        if (data[i][0] === 'B；支払合計') {
          updatedTotalRows.push(i);
          break;
        }
      }
    }
    
    // 収支情報セクション
    let profitSectionIndex = -1;
    for (let i = expenseSectionIndex + 2; i < data.length; i++) {
      if (data[i][0] === '◆ 収支情報') {
        profitSectionIndex = i;
        break;
      }
    }
    
    if (profitSectionIndex > 0) {
      updatedSectionHeaders.push(profitSectionIndex);
      updatedHeaderRows.push(profitSectionIndex + 1);
      
      // 上席チェック行を探す
      for (let i = profitSectionIndex + 2; i < data.length; i++) {
        if (data[i][0] === '上席チェック') {
          updatedHeaderRows.push(i);
          break;
        }
      }
    }
    
    // マージセルの更新
    const updatedMerges = [];
    
    // 基本情報セクションのマージセル
    updatedMerges.push({row: basicSectionIndex, col: 0, rowspan: 1, colspan: 8});
    updatedMerges.push({row: basicSectionIndex + 1, col: 0, rowspan: 1, colspan: 2});
    
    // 入金情報セクションのマージセル
    if (paymentSectionIndex > 0) {
      updatedMerges.push({row: paymentSectionIndex, col: 0, rowspan: 1, colspan: 8});
      
      // 入金合計行
      const paymentTotalIndex = updatedTotalRows.find(idx => data[idx][0] === 'A；入金合計');
      if (paymentTotalIndex) {
        updatedMerges.push({row: paymentTotalIndex, col: 0, rowspan: 1, colspan: 7});
      }
    }
    
    // 支払情報セクションのマージセル
    if (expenseSectionIndex > 0) {
      updatedMerges.push({row: expenseSectionIndex, col: 0, rowspan: 1, colspan: 8});
      
      // 支払合計行
      const expenseTotalIndex = updatedTotalRows.find(idx => data[idx][0] === 'B；支払合計');
      if (expenseTotalIndex) {
        updatedMerges.push({row: expenseTotalIndex, col: 0, rowspan: 1, colspan: 7});
      }
    }
    
    // 収支情報セクションのマージセル
    if (profitSectionIndex > 0) {
      updatedMerges.push({row: profitSectionIndex, col: 0, rowspan: 1, colspan: 8});
      
      // 上席チェック行
      for (let i = profitSectionIndex + 2; i < data.length; i++) {
        if (data[i][0] === '上席チェック') {
          updatedMerges.push({row: i, col: 0, rowspan: 1, colspan: 2});
          break;
        }
      }
    }
    
    // スタイル情報を更新
    SECTION_STYLES.sectionHeaders = updatedSectionHeaders;
    SECTION_STYLES.headerRows = updatedHeaderRows;
    SECTION_STYLES.totalRows = updatedTotalRows;
    
    // マージセルと設定を更新
    this.hot.updateSettings({
      mergeCells: updatedMerges,
      cells: function(row, col) {
        const cellProperties = {};
        
        // セクション見出し行のスタイル
        if (updatedSectionHeaders.includes(row)) {
          cellProperties.renderer = SpreadsheetManager.sectionHeaderRenderer;
          cellProperties.readOnly = true;
        }
        // ヘッダー行のスタイル
        else if (updatedHeaderRows.includes(row)) {
          cellProperties.renderer = SpreadsheetManager.headerRenderer;
        }
        // 合計行のスタイル
        else if (updatedTotalRows.includes(row)) {
          cellProperties.renderer = SpreadsheetManager.totalRenderer;
        }
        
        // 1列目（項目名列）の設定
        if (col === 0) {
          cellProperties.wordWrap = true;
        }
        
        return cellProperties;
      }
    });
    
    // 再レンダリング
    this.hot.render();
  },
  
  /**
   * カスタムセルレンダラー：セクション見出しセル
   */
  sectionHeaderRenderer: function(instance, td, row, col, prop, value, cellProperties) {
    Handsontable.renderers.TextRenderer.apply(this, arguments);
    td.style.backgroundColor = '#4472C4';
    td.style.color = 'white';
    td.style.fontWeight = 'bold';
    td.style.fontSize = '14px';
    td.style.padding = '6px';
  },
  
  /**
   * カスタムセルレンダラー：ヘッダーセル
   */
  headerRenderer: function(instance, td, row, col, prop, value, cellProperties) {
    Handsontable.renderers.TextRenderer.apply(this, arguments);
    td.style.backgroundColor = '#f3f2f1';
    td.style.fontWeight = 'bold';
    td.style.textAlign = col === 0 ? 'left' : 'center';
  },
  
  /**
   * カスタムセルレンダラー：合計セル
   */
  totalRenderer: function(instance, td, row, col, prop, value, cellProperties) {
    Handsontable.renderers.TextRenderer.apply(this, arguments);
    td.style.backgroundColor = '#e6f2ff';
    td.style.fontWeight = 'bold';
  },
  
  /**
   * カスタムセルレンダラー：区切り行
   */
  spacerRenderer: function(instance, td, row, col, prop, value, cellProperties) {
    Handsontable.renderers.TextRenderer.apply(this, arguments);
    td.style.borderBottom = '1px solid #ccc';
    td.style.height = '20px';
  },
  
  /**
   * データをスプレッドシートにロード
   * @param {Array} data - ロードするデータ
   */
  loadData: function(data) {
    if (!this.hot) {
      console.warn('Handsontableインスタンスがないため、初期化します');
      this.initialize();
      if (!this.hot) {
        console.error('Handsontableの初期化に失敗しました');
        return;
      }
    }
    
    try {
      // データが30行未満の場合、30行になるよう空行を追加
      if (data.length < 30) {
        const emptyRows = Array(30 - data.length).fill().map(() => Array(8).fill(''));
        data = [...data, ...emptyRows];
      }
      
      console.log('データをロードします');
      this.hot.loadData(data);
      
      // データロード後、スクロールを一番上に戻す
      setTimeout(() => {
        const container = document.querySelector('.spreadsheet-container');
        if (container) {
          container.scrollTop = 0;
        }
        
        // 再レンダリング
        this.hot.render();
      }, 100);
    } catch (error) {
      console.error('データロード中にエラーが発生しました:', error);
    }
  },
  
  /**
   * 現在のスプレッドシートのデータを取得
   * @return {Array} スプレッドシートのデータ
   */
  getData: function() {
    if (!this.hot) {
      console.warn('Handsontableインスタンスがありません');
      return [];
    }
    return this.hot.getData();
  },
  
  /**
   * セルの値を設定
   * @param {number} row - 行インデックス
   * @param {number} col - 列インデックス
   * @param {*} value - 設定する値
   */
  setDataAtCell: function(row, col, value) {
    if (!this.hot) {
      console.warn('Handsontableインスタンスがありません');
      return;
    }
    this.hot.setDataAtCell(row, col, value);
  },
  
  /**
   * 統一テンプレートをロード
   */
  loadUnifiedTemplate: function() {
    console.log('テンプレートをロードします');
    this.loadData(UNIFIED_TEMPLATE);
  },
  
  /**
   * Excelファイルをインポート
   * @param {File} file - インポートするExcelファイル
   * @param {Function} callback - インポート完了時のコールバック
   */
  importFromExcel: function(file, callback) {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        
        // 最初のシートのデータを取得
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
        
        // データをHandsontableに読み込む
        this.loadData(jsonData);
        
        if (callback) callback(true);
      } catch (error) {
        console.error('Import error:', error);
        if (callback) callback(false, error.message);
      }
    };
    reader.readAsArrayBuffer(file);
  },
  
  /**
   * 現在のデータをExcelファイルとしてエクスポート
   * @param {string} filename - 出力ファイル名
   */
  exportToExcel: function(filename) {
    try {
      // 現在のシートデータを取得
      const data = this.getData();
      
      // ワークシートを作成
      const ws = XLSX.utils.aoa_to_sheet(data);
      
      // 列幅の設定
      ws['!cols'] = COLUMN_WIDTHS.map(width => ({ width: width / 8 }));
      
      // セル結合の設定
      const merges = this.hot.getSettings().mergeCells || [];
      ws['!merges'] = merges.map(merge => ({ 
        s: { r: merge.row, c: merge.col },
        e: { r: merge.row + merge.rowspan - 1, c: merge.col + merge.colspan - 1 }
      }));
      
      // 新しいワークブックを作成してシートを追加
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'カルテデータ');
      
      // Excelファイルとして出力
      XLSX.writeFile(wb, filename);
      return true;
    } catch (error) {
      console.error('Export error:', error);
      return false;
    }
  }
};
