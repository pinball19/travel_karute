/**
 * アプリケーションのメインエントリーポイント
 */
document.addEventListener('DOMContentLoaded', function() {
  console.log('DOMContentLoaded: アプリケーションを初期化します');
  
  // DOM要素
  const saveButton = document.getElementById('save-button');
  const newButton = document.getElementById('new-button');
  const openButton = document.getElementById('open-button');
  const importButton = document.getElementById('import-button');
  const exportButton = document.getElementById('export-button');
  const printButton = document.getElementById('print-button');
  const fileImport = document.getElementById('file-import');
  const closeModal = document.querySelector('.close-modal');
  const karteListModal = document.getElementById('karte-list-modal');
  
  // アプリケーションの状態変数
  let isInitialized = false;
  
  // 初期化処理
  function initialize() {
    if (isInitialized) return;
    
    try {
      console.log('スプレッドシートを初期化します');
      
      // スプレッドシートの初期化
      SpreadsheetManager.initialize();
      
      isInitialized = true;
      console.log('初期化完了');
    } catch (error) {
      console.error('初期化中にエラーが発生しました:', error);
      
      // 500ms後に再試行
      setTimeout(() => {
        if (!isInitialized) {
          console.log('初期化を再試行します');
          initialize();
        }
      }, 500);
    }
  }
  
  // 初期化を実行
  initialize();
  
  // 初期化が間違いなく実行されるようにする
  setTimeout(() => {
    if (!isInitialized) {
      console.log('初期化の再確認');
      initialize();
    }
  }, 1000);
  
  // 保存ボタンのイベント
  saveButton.addEventListener('click', () => {
    try {
      console.log('カルテを保存します');
      KarteManager.saveKarte();
    } catch (error) {
      console.error('保存中にエラーが発生しました:', error);
      alert(`保存中にエラーが発生しました: ${error.message}`);
    }
  });
  
  // 新規ボタンのイベント
  newButton.addEventListener('click', () => {
    try {
      console.log('新規カルテを作成します');
      KarteManager.createNewKarte();
    } catch (error) {
      console.error('新規作成中にエラーが発生しました:', error);
      alert(`新規作成中にエラーが発生しました: ${error.message}`);
    }
  });
  
  // 開くボタンのイベント
  openButton.addEventListener('click', () => {
    try {
      console.log('カルテ一覧を表示します');
      KarteManager.loadKarteList();
      karteListModal.style.display = 'block';
    } catch (error) {
      console.error('カルテ一覧の読み込み中にエラーが発生しました:', error);
      alert(`カルテ一覧の読み込み中にエラーが発生しました: ${error.message}`);
    }
  });
  
  // モーダルを閉じるボタンのイベント
  closeModal.addEventListener('click', () => {
    karteListModal.style.display = 'none';
  });
  
  // モーダル外をクリックしたら閉じる
  window.addEventListener('click', (event) => {
    if (event.target === karteListModal) {
      karteListModal.style.display = 'none';
    }
  });
  
  // インポートボタンのイベント
  importButton.addEventListener('click', () => {
    fileImport.click();
  });
  
  // ファイル選択イベント
  fileImport.addEventListener('change', (event) => {
    const file = event.target.files[0];
    if (!file) return;
    
    try {
      console.log('Excelファイルをインポートします');
      SpreadsheetManager.importFromExcel(file, (success, errorMessage) => {
        if (success) {
          // インポート後は未保存状態
          KarteManager.currentKarteId = null;
          document.getElementById('current-karte-id').textContent = '新規カルテ';
          document.getElementById('last-saved').textContent = '保存されていません';
          
          alert('Excelファイルをインポートしました');
        } else {
          alert(`インポートエラー: ${errorMessage}`);
        }
      });
    } catch (error) {
      console.error('インポート中にエラーが発生しました:', error);
      alert(`インポート中にエラーが発生しました: ${error.message}`);
    }
    
    // input要素をリセット
    event.target.value = '';
  });
  
  // エクスポートボタンのイベント
  exportButton.addEventListener('click', () => {
    try {
      console.log('Excelファイルをエクスポートします');
      const filename = `カルテ_${new Date().toISOString().slice(0, 10)}.xlsx`;
      const success = SpreadsheetManager.exportToExcel(filename);
      
      if (!success) {
        alert('エクスポート中にエラーが発生しました');
      }
    } catch (error) {
      console.error('エクスポート中にエラーが発生しました:', error);
      alert(`エクスポート中にエラーが発生しました: ${error.message}`);
    }
  });
  
  // 印刷ボタンのイベント
  printButton.addEventListener('click', () => {
    try {
      console.log('印刷します');
      window.print();
    } catch (error) {
      console.error('印刷中にエラーが発生しました:', error);
      alert(`印刷中にエラーが発生しました: ${error.message}`);
    }
  });
  
  // オンライン状態の監視
  const connectionStatus = document.getElementById('connection-status');
  
  window.addEventListener('online', () => {
    connectionStatus.textContent = 'オンライン';
    connectionStatus.style.color = 'green';
  });
  
  window.addEventListener('offline', () => {
    connectionStatus.textContent = 'オフライン';
    connectionStatus.style.color = 'red';
  });
  
  // テーブルが表示されない問題を修正するためのフォールバック
  window.addEventListener('load', function() {
    setTimeout(function() {
      const hotContainer = document.getElementById('hot-container');
      if (!hotContainer.querySelector('.handsontable') || 
          !document.querySelector('.handsontable .ht_master .wtHolder')) {
        console.warn('Handsontableが見つかりません。再初期化します。');
        SpreadsheetManager.initialize();
      }
    }, 1500);
  });
});
