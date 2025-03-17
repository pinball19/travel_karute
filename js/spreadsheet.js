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
    
    // 基本設定
    const options = {
      ...APP_CONFIG.HOT_OPTIONS,
      data: UNIFIED_TEMPLATE,
      mergeCells: UNIFIED_MERGED_CELLS,
      
      // カラム幅を設定
      colWidths: COLUMN_WIDTHS,
      
      // 自動行高さ調整を有効化
      autoRowSize: {
        syncLimit: 1000
      },
      
      // ワードラップを有効化（テキスト折り返し）
      wordWrap: true,
      
      // 最小表示行数を設定（30行以上表示）
      minRows: 30,
      minSpareRows: 5, // 空の行を追加して下部に余白を確保
      
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
          cellProperties.wordWrap = true; // テキスト折り返しを特に有効化
        }
        
        return cellProperties;
      },
      
      // 行ヘッダーを常に表示
      fixedRowsTop: 0,
      fixedColumnsLeft: 0,
      
      // スクロール設定
      viewportRowRenderingOffset: 30, // より多くの行を事前にレンダリング
    };
    
    // Handsontableインスタンスを作成
    this.hot = new Handsontable(container, options);
    
    // ウィンドウのリサイズに合わせてHandsontableをリサイズ
    window.addEventListener('resize', () => {
      this.hot.updateSettings({
        width: container.offsetWidth,
        height: container.offsetHeight
      });
    });
    
    // 初期化後、全体を表示するためにセルの再レンダリングを実行
    setTimeout(() => {
      this.hot.render();
      
      // スクロールを一番上に戻す
      container.scrollTop = 0;
    }, 100);
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
    if (!this.hot) return;
    
    // データが30行未満の場合、30行になるよう空行を追加
    if (data.length < 30) {
      const emptyRows = Array(30 - data.length).fill().map(() => Array(8).fill(''));
      data = [...data, ...emptyRows];
    }
    
    this.hot.loadData(data);
    
    // データロード後、スクロールを一番上に戻す
    setTimeout(() => {
      document.querySelector('.spreadsheet-container').scrollTop = 0;
    }, 100);
  },
  
  /**
   * 現在のスプレッドシートのデータを取得
   * @return {Array} スプレッドシートのデータ
   */
  getData: function() {
    if (!this.hot) return [];
    return this.hot.getData();
  },
  
  /**
   * セルの値を設定
   * @param {number} row - 行インデックス
   * @param {number} col - 列インデックス
   * @param {*} value - 設定する値
   */
  setDataAtCell: function(row, col, value) {
    if (!this.hot) return;
    this.hot.setDataAtCell(row, col, value);
  },
  
  /**
   * 統一テンプレートをロード
   */
  loadUnifiedTemplate: function() {
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
