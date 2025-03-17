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
      data: TEMPLATES[0],
      mergeCells: BASIC_INFO_MERGED_CELLS,
      cells: function(row, col) {
        const cellProperties = {};
        
        // ヘッダー行のスタイル
        if (row === 0 || row === 1 || row === 4 || row === 9 || row === 14) {
          cellProperties.renderer = SpreadsheetManager.headerRenderer;
        }
        
        // 合計行のスタイル
        if (row === 8 || row === 13) {
          cellProperties.renderer = SpreadsheetManager.totalRenderer;
        }
        
        return cellProperties;
      }
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
  },
  
  /**
   * カスタムセルレンダラー：ヘッダーセル
   */
  headerRenderer: function(instance, td, row, col, prop, value, cellProperties) {
    Handsontable.renderers.TextRenderer.apply(this, arguments);
    td.style.backgroundColor = '#f3f2f1';
    td.style.fontWeight = 'bold';
    td.style.textAlign = 'center';
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
   * データをスプレッドシートにロード
   * @param {Array} data - ロードするデータ
   */
  loadData: function(data) {
    if (!this.hot) return;
    this.hot.loadData(data);
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
   * シートのタイプに応じたテンプレートをロード
   * @param {number} sheetIndex - シートのインデックス
   */
  loadTemplateBySheetIndex: function(sheetIndex) {
    this.loadData(TEMPLATES[sheetIndex]);
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