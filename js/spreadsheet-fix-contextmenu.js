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
    
    // コンテナのサイズを確認
    if (container.offsetHeight === 0) {
      console.warn('hot-containerの高さがゼロです。500ms後に再試行します');
      setTimeout(() => this.initialize(), 500);
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
        contextMenu: true, // シンプルに true を指定
        manualColumnResize: true,
        manualRowResize: true,
        licenseKey: 'non-commercial-and-evaluation',
        mergeCells: UNIFIED_MERGED_CELLS,
        
        // カラム幅を設定
        colWidths: COLUMN_WIDTHS,
        
        // 表示設定 - スクロールバーを無効化
        width: '100%',
        height: '100%',
        
        // テキスト折り返し
        wordWrap: true,
        
        // 最小表示行数
        minRows: 30,
        minSpareRows: 5,
        
        // スクロールバー設定
        viewportColumnRenderingOffset: 30, // 表示範囲外の列も描画
        viewportRowRenderingOffset: 30,    // 表示範囲外の行も描画
        
        // モバイル対応設定
        outsideClickDeselects: false, // 外部クリックで選択解除しない
        fragmentSelection: false,     // モバイルでのセル選択を改善
        
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
        
        // 行の高さ自動調整
        autoRowSize: {
          syncLimit: 1000
        },
        
        // データ変更後に合計を計算
        afterChange: function(changes, source) {
          if (source === 'edit' || source === 'paste') {
            // 少し遅延させて実行（データが確実に更新された後）
            setTimeout(() => {
              SpreadsheetManager.calculateSums();
            }, 100);
          }
        }
      };
      
      console.log('Handsontableを初期化します');
      // Handsontableインスタンスを作成
      this.hot = new Handsontable(container, options);
      console.log('Handsontableの初期化完了');
      
      // レンダリングを確実に行う
      setTimeout(() => {
        if (this.hot) {
          console.log('Handsontableを再レンダリングします');
          this.hot.render();
          
          // スクロールバーの設定を上書き
          const wtHolders = document.querySelectorAll('.wtHolder');
          wtHolders.forEach(holder => {
            holder.style.overflowX = 'visible';
            holder.style.overflowY = 'visible';
          });
          
          container.scrollTop = 0;
          
          // 初期化時に合計を計算
          this.calculateSums();
        }
      }, 200);
      
      // モバイルデバイスで操作しやすいように調整
      this.adjustForMobile();
      
    } catch (error) {
      console.error('Handsontableの初期化中にエラーが発生しました:', error);
    }
    
    // ウィンドウのリサイズに合わせてHandsontableをリサイズ
    window.addEventListener('resize', () => {
      if (this.hot) {
        console.log('リサイズによりHandsontableを更新します');
        this.hot.updateSettings({
          width: '100%',
          height: '100%'
        });
        this.hot.render();
        
        // モバイルデバイスの場合は追加の調整
        this.adjustForMobile();
      }
    });
  },
  
  /**
   * 入金合計と支払合計を計算する
   */
  calculateSums: function() {
    if (!this.hot) return;
    
    try {
      // DOM要素の確認
      if (!this.hot.rootElement || !this.hot.rootElement.offsetHeight) {
        console.warn('Handsontableの要素が正しく初期化されていません。計算をスキップします');
        return;
      }
      
      // データを取得
      const data = this.hot.getData();
      if (!data || !Array.isArray(data)) {
        console.warn('データが正しく取得できませんでした');
        return;
      }
      
      // 入金情報の合計を計算（8行目から11行目、4列目の値を合計）
      let paymentSum = 0;
      for (let i = 8; i <= 11; i++) {
        if (data[i] && data[i][4]) {
          const value = parseFloat(data[i][4]);
          if (!isNaN(value)) {
            paymentSum += value;
          }
        }
      }
      
      // 支払情報の合計を計算（16行目から19行目、5列目の値を合計）
      let expenseSum = 0;
      for (let i = 16; i <= 19; i++) {
        if (data[i] && data[i][5]) {
          const value = parseFloat(data[i][5]);
          if (!isNaN(value)) {
            expenseSum += value;
          }
        }
      }
      
      // 安全にセルに値を設定する関数
      const safeSetDataAtCell = (row, col, value) => {
        try {
          if (this.hot && this.hot.setDataAtCell && typeof row === 'number' && typeof col === 'number') {
            this.hot.setDataAtCell(row, col, value);
          }
        } catch (error) {
          console.warn(`セル(${row},${col})への設定中にエラーが発生しました:`, error);
        }
      };
      
      // 入金合計を設定（12行目、4列目）
      safeSetDataAtCell(12, 4, paymentSum > 0 ? paymentSum : '');
      
      // 支払合計を設定（20行目、5列目）
      safeSetDataAtCell(20, 5, expenseSum > 0 ? expenseSum : '');
      
      // 収支情報を更新（24行目）
      if (paymentSum > 0 || expenseSum > 0) {
        // 利益額（収入 - 支出）
        const profit = paymentSum - expenseSum;
        safeSetDataAtCell(24, 2, profit !== 0 ? profit : '');
        
        // 利益率（利益額 / 収入 * 100）
        if (paymentSum > 0) {
          const profitRate = (profit / paymentSum * 100).toFixed(1);
          safeSetDataAtCell(24, 1, profitRate + '%');
        }
        
        // 旅行総額と支払総額
        safeSetDataAtCell(24, 4, paymentSum > 0 ? paymentSum : '');
        safeSetDataAtCell(24, 5, expenseSum > 0 ? expenseSum : '');
        
        // 人数を取得（4行目、5列目）
        if (data[3] && data[3][5]) {
          const persons = parseFloat(data[3][5]);
          if (!isNaN(persons) && persons > 0) {
            // 一人粗利を計算
            const perPersonProfit = (profit / persons).toFixed(0);
            safeSetDataAtCell(24, 3, perPersonProfit !== '0' ? perPersonProfit : '');
            
            // 人数を収支情報にもコピー
            safeSetDataAtCell(24, 6, persons);
          }
        }
      }
    } catch (error) {
      console.error('合計計算中にエラーが発生しました:', error);
    }
  },
  
  /**
   * モバイルデバイス向けの調整
   */
  adjustForMobile: function() {
    try {
      if (window.innerWidth <= 768) { // スマホサイズの場合
        // ヘッダーの高さを調整
        const colHeaders = document.querySelectorAll('.ht_clone_top .htCore th');
        colHeaders.forEach(th => {
          th.style.height = '30px';
          th.style.padding = '2px';
          th.style.fontSize = '12px';
        });
        
        // 行ヘッダーの幅を調整
        const rowHeaders = document.querySelectorAll('.ht_clone_left .htCore th');
        rowHeaders.forEach(th => {
          th.style.width = '30px';
          th.style.padding = '2px';
          th.style.fontSize = '12px';
        });
        
        // セルのサイズを調整
        const cells = document.querySelectorAll('.htCore td');
        cells.forEach(td => {
          td.style.padding = '4px';
          td.style.fontSize = '12px';
        });
      }
    } catch (error) {
      console.error('モバイル調整中にエラーが発生しました:', error);
    }
  },
  
  /**
   * 行の操作（挿入/削除）後に背景色などを更新
   */
  updateAfterRowOperation: function() {
    if (!this.hot) return;
    
    try {
      // データを取得
      const data = this.hot.getData();
      if (!data || !Array.isArray(data)) {
        console.warn('データが正しく取得できませんでした');
        return;
      }
      
      // セクションヘッダーの位置を探し、スタイル設定を更新
      let updatedSectionHeaders = [];
      let updatedHeaderRows = [];
      let updatedTotalRows = [];
      let updatedSpacerRows = [];
      
      // 基本情報セクション
      const basicSectionIndex = 0;
      updatedSectionHeaders.push(basicSectionIndex);
      updatedHeaderRows.push(basicSectionIndex + 1);
      updatedHeaderRows.push(basicSectionIndex + 2);
      updatedHeaderRows.push(basicSectionIndex + 3); // 4行目も追加
      updatedHeaderRows.push(basicSectionIndex + 4); // 5行目も追加
      updatedSpacerRows.push(basicSectionIndex + 5);
      
      // 入金情報セクション
      let paymentSectionIndex = -1;
      for (let i = basicSectionIndex + 3; i < data.length; i++) {
        if (data[i] && data[i][0] === '◆ 入金情報') {
          paymentSectionIndex = i;
          break;
        }
      }
      
      if (paymentSectionIndex > 0) {
        updatedSectionHeaders.push(paymentSectionIndex);
        updatedHeaderRows.push(paymentSectionIndex + 1);
        updatedSpacerRows.push(paymentSectionIndex + 6);
        
        // 入金合計行を探す
        for (let i = paymentSectionIndex + 2; i < data.length; i++) {
          if (data[i] && data[i][0] === 'A；入金合計') {
            updatedTotalRows.push(i);
            break;
          }
        }
      }
      
      // 支払情報セクション
      let expenseSectionIndex = -1;
      for (let i = (paymentSectionIndex > 0 ? paymentSectionIndex + 2 : basicSectionIndex + 6); i < data.length; i++) {
        if (data[i] && data[i][0] === '◆ 支払情報') {
          expenseSectionIndex = i;
          break;
        }
      }
      
      if (expenseSectionIndex > 0) {
        updatedSectionHeaders.push(expenseSectionIndex);
        updatedHeaderRows.push(expenseSectionIndex + 1);
        updatedSpacerRows.push(expenseSectionIndex + 6);
        
        // 支払合計行を探す
        for (let i = expenseSectionIndex + 2; i < data.length; i++) {
          if (data[i] && data[i][0] === 'B；支払合計') {
            updatedTotalRows.push(i);
            break;
          }
        }
      }
      
      // 収支情報セクション
      let profitSectionIndex = -1;
      for (let i = (expenseSectionIndex > 0 ? expenseSectionIndex + 2 : basicSectionIndex + 12); i < data.length; i++) {
        if (data[i] && data[i][0] === '◆ 収支情報') {
          profitSectionIndex = i;
          break;
        }
      }
      
      if (profitSectionIndex > 0) {
        updatedSectionHeaders.push(profitSectionIndex);
        updatedHeaderRows.push(profitSectionIndex + 1);
        
        // 上席チェック行を探す
        for (let i = profitSectionIndex + 2; i < data.length; i++) {
          if (data[i] && data[i][0] === '上席チェック') {
            updatedHeaderRows.push(i);
            break;
          }
        }
        
        // 最後の数行を区切り行に
        for (let i = profitSectionIndex + 5; i < Math.min(profitSectionIndex + 10, data.length); i++) {
          updatedSpacerRows.push(i);
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
        const paymentTotalRow = updatedTotalRows.find(idx => data[idx] && data[idx][0] === 'A；入金合計');
        if (paymentTotalRow) {
          updatedMerges.push({row: paymentTotalRow, col: 0, rowspan: 1, colspan: 7});
        }
      }
      
      // 支払情報セクションのマージセル
      if (expenseSectionIndex > 0) {
        updatedMerges.push({row: expenseSectionIndex, col: 0, rowspan: 1, colspan: 8});
        
        // 支払合計行
        const expenseTotalRow = updatedTotalRows.find(idx => data[idx] && data[idx][0] === 'B；支払合計');
        if (expenseTotalRow) {
          updatedMerges.push({row: expenseTotalRow, col: 0, rowspan: 1, colspan: 7});
        }
      }
      
      // 収支情報セクションのマージセル
      if (profitSectionIndex > 0) {
        updatedMerges.push({row: profitSectionIndex, col: 0, rowspan: 1, colspan: 8});
        
        // 上席チェック行
        for (let i = profitSectionIndex + 2; i < data.length; i++) {
          if (data[i] && data[i][0] === '上席チェック') {
            updatedMerges.push({row: i, col: 0, rowspan: 1, colspan: 2});
            break;
          }
        }
      }
      
      // スタイル情報を更新
      SECTION_STYLES.sectionHeaders = updatedSectionHeaders;
      SECTION_STYLES.headerRows = updatedHeaderRows;
      SECTION_STYLES.totalRows = updatedTotalRows;
      SECTION_STYLES.spacerRows = updatedSpacerRows;
      
      // マージセルと設定を更新
      try {
        if (this.hot && this.hot.updateSettings) {
          this.hot.updateSettings({
            mergeCells: updatedMerges
          });
          
          // 再レンダリング
          this.hot.render();
        }
      } catch (updateError) {
        console.error('設定更新中にエラーが発生しました:', updateError);
      }
      
      // 合計を再計算
      this.calculateSums();
    } catch (error) {
      console.error('行操作後の更新中にエラーが発生しました:', error);
    }
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
      // データの検証
      if (!data || !Array.isArray(data)) {
        console.error('無効なデータが提供されました');
        data = UNIFIED_TEMPLATE; // デフォルトテンプレートを使用
      }
      
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
        if (this.hot && this.hot.render) {
          this.hot.render();
        }
        
        // スクロールバーの設定を上書き
        const wtHolders = document.querySelectorAll('.wtHolder');
        wtHolders.forEach(holder => {
          holder.style.overflowX = 'visible';
          holder.style.overflowY = 'visible';
        });
        
        // 合計を計算
        this.calculateSums();
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
    try {
      return this.hot.getData();
    } catch (error) {
      console.error('データ取得中にエラーが発生しました:', error);
      return [];
    }
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
    try {
      this.hot.setDataAtCell(row, col, value);
    } catch (error) {
      console.error(`セル(${row},${col})への設定中にエラーが発生しました:`, error);
    }
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
