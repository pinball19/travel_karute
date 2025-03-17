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
        colHeaders: false,
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
        
        // セルのカスタマイズ
        cells: function(row, col) {
          const cellProperties = {};
          
          // セクション見出し行のスタイル
          if (SECTION_STYLES.sectionHeaders.includes(row)) {
            cellProperties.renderer = SpreadsheetManager.sectionHeaderRenderer;
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
      ws['!merges'] = UNIFIED_MERGED_CELLS.map(merge => ({ 
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
