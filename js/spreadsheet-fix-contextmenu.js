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
    
    // デバッグ情報を追加
    console.log('コンテナ要素:', container);
    
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
        contextMenu: {
          items: {
            'row_above': { name: '上に行を挿入' },
            'row_below': { name: '下に行を挿入' },
            'remove_row': { name: '行を削除' },
            'separator1': Handsontable.plugins.ContextMenu.SEPARATOR,
            'copy': { name: 'コピー' },
            'cut': { name: '切り取り' },
            'separator2': Handsontable.plugins.ContextMenu.SEPARATOR,
            'alignment': {
              name: '配置',
              submenu: {
                items: [
                  { key: 'alignment:left', name: '左揃え' },
                  { key: 'alignment:center', name: '中央揃え' },
                  { key: 'alignment:right', name: '右揃え' }
                ]
              }
            }
          }
        },
        manualColumnResize: true,
        manualRowResize: false, // 行の高さを手動変更不可に設定
        licenseKey: 'non-commercial-and-evaluation',
        mergeCells: UNIFIED_MERGED_CELLS,
        
        // カラム幅を設定
        colWidths: COLUMN_WIDTHS,
        
        // 行の高さを固定
        rowHeights: 25,
        
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
        fragmentSelection: 'cell',    // モバイルでのセル選択を改善
        
        // 固定列数
        fixedColumnsLeft: 0, // 左側の固定列を無効化
        
        // ヘッダー固定
        fixedRowsTop: 0, // ヘッダー行を固定しない
        
        // 挿入・削除時のフック
        afterCreateRow: function(index, amount) {
          console.log(`${amount}行が位置${index}に挿入されました`);
          SpreadsheetManager.updateAfterRowOperation();
        },
        
        afterRemoveRow: function(index, amount) {
          console.log(`${amount}行が位置${index}から削除されました`);
          SpreadsheetManager.updateAfterRowOperation();
        },
        
        // 描画フック
        afterRender: function() {
          // レンダリング後に行の高さを調整
          SpreadsheetManager.adjustRowHeights();
          // スタイルを明示的に適用
          SpreadsheetManager.applyCustomStyles();
        },
        
        // 行の高さ自動調整を無効化
        autoRowSize: false,
        
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
          
          // スタイルを明示的に適用
          this.applyCustomStyles();
          
          // 列と行の位置を調整
          this.fixRowAlignment();
          
          container.scrollTop = 0;
        }
      }, 500); // タイムアウトを延長
      
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
        
        // 列と行の位置を調整
        this.fixRowAlignment();
        
        // スタイルを明示的に適用
        this.applyCustomStyles();
        
        // モバイルデバイスの場合は追加の調整
        this.adjustForMobile();
      }
    });
  },
  
  /**
   * スタイルを明示的に適用する関数
   */
  applyCustomStyles: function() {
    // テーブル要素を確認
    const tableElements = document.querySelectorAll('.handsontable .htCore');
    console.log('テーブル要素数:', tableElements.length);
    
    // 各要素にスタイルを適用
    tableElements.forEach(table => {
      table.style.borderCollapse = 'collapse';
      table.style.width = '100%';
    });
    
    // セルに明示的にスタイルを適用
    const cells = document.querySelectorAll('.handsontable .htCore td');
    cells.forEach(cell => {
      cell.style.border = '1px solid #e0e0e0';
      cell.style.padding = '4px';
      cell.style.height = '25px';
    });
    
    // ヘッダーセルにスタイルを適用
    const headerCells = document.querySelectorAll('.handsontable .htCore th');
    headerCells.forEach(th => {
      th.style.border = '1px solid #ccc';
      th.style.background = '#f3f2f1';
    });
    
    // スクロールバー要素に明示的にスタイルを適用
    const scrollbars = document.querySelectorAll('.handsontable .wtHolder');
    scrollbars.forEach(holder => {
      holder.style.overflow = 'auto';
    });
  },
  
  /**
   * 行の高さを調整する
   */
  adjustRowHeights: function() {
    if (!this.hot) return;
    
    // セクションヘッダーとその他の行で高さを分ける
    const rowCount = this.hot.countRows();
    const rowHeights = Array(rowCount).fill(25); // デフォルトは25px
    
    // セクションヘッダーの行は高さを大きく
    SECTION_STYLES.sectionHeaders.forEach(rowIndex => {
      if (rowIndex < rowCount) {
        rowHeights[rowIndex] = 30;
      }
    });
    
    // 行の高さを設定
    this.hot.updateSettings({
      rowHeights: rowHeights
    });
    
    // 明示的に行の高さをDOM要素に適用
    setTimeout(() => {
      const rows = document.querySelectorAll('.handsontable .htCore tr');
      rows.forEach((row, index) => {
        if (index < rowHeights.length) {
          row.style.height = `${rowHeights[index]}px`;
        }
      });
    }, 100);
  },
  
  /**
   * 列と行の位置ずれを修正する
   */
  fixRowAlignment: function() {
    if (!this.hot) return;
    
    // すべての行の高さを同期
    const rowCount = this.hot.countRows();
    
    // 左側固定列と本体の行の高さを同期
    for (let i = 0; i < rowCount; i++) {
      const mainRow = document.querySelector(`.ht_master .htCore tr:nth-child(${i + 1})`);
      const leftRow = document.querySelector(`.ht_clone_left .htCore tr:nth-child(${i + 1})`);
      
      if (mainRow && leftRow) {
        const height = `${mainRow.offsetHeight}px`;
        mainRow.style.height = height;
        leftRow.style.height = height;
      }
    }
    
    // テーブル全体を描画しなおす
    this.hot.render();
  },
  
  /**
   * モバイルデバイス向けの調整
   */
  adjustForMobile: function() {
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
  },
  
  /**
   * 行の操作（挿入/削除）後に背景色などを更新
   */
  updateAfterRowOperation: function() {
    if (!this.hot) return;
    
    try {
      // データを取得
      const data = this.hot.getData();
      
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
      updatedSpacerRows.push(basicSectionIndex + 5);
      
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
        updatedSpacerRows.push(paymentSectionIndex + 6);
        
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
      for (let i = (paymentSectionIndex > 0 ? paymentSectionIndex + 2 : basicSectionIndex + 6); i < data.length; i++) {
        if (data[i][0] === '◆ 支払情報') {
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
          if (data[i][0] === 'B；支払合計') {
            updatedTotalRows.push(i);
            break;
          }
        }
      }
      
      // 収支情報セクション
      let profitSectionIndex = -1;
      for (let i = (expenseSectionIndex > 0 ? expenseSectionIndex + 2 : basicSectionIndex + 12); i < data.length; i++) {
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
