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
  let lastChangeTime = 0;
  let autoSaveTimer = null;
  
  // 初期化処理
  function initialize() {
    if (isInitialized) return;
    
    try {
      console.log('スプレッドシートを初期化します');
      
      // スプレッドシートの初期化
      SpreadsheetManager.initialize();
      
      // 自動保存の設定
      setupAutoSave();
      
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
  
  // 自動保存の設定
  function setupAutoSave() {
    if (SpreadsheetManager.hot) {
      // データ変更時のイベント
      SpreadsheetManager.hot.addHook('afterChange', function(changes, source) {
        if (source === 'edit' || source === 'paste') {
          lastChangeTime = Date.now();
          
          // 入金合計と支払合計を自動計算
          SpreadsheetManager.calculateSums();
          
          // 既存のタイマーをクリア
          if (autoSaveTimer) {
            clearTimeout(autoSaveTimer);
          }
          
          // 3秒後に自動保存
          autoSaveTimer = setTimeout(() => {
            console.log('自動保存を実行します');
            KarteManager.saveKarte();
          }, 3000);
        }
      });
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