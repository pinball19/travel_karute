/**
 * アプリケーションのメインエントリーポイント
 */
document.addEventListener('DOMContentLoaded', function() {
  // DOM要素
  const saveButton = document.getElementById('save-button');
  const newButton = document.getElementById('new-button');
  const openButton = document.getElementById('open-button');
  const importButton = document.getElementById('import-button');
  const exportButton = document.getElementById('export-button');
  const printButton = document.getElementById('print-button');
  const fileImport = document.getElementById('file-import');
  const sheetTabs = document.querySelectorAll('.sheet-tab');
  const closeModal = document.querySelector('.close-modal');
  const karteListModal = document.getElementById('karte-list-modal');
  
  // スプレッドシートの初期化
  SpreadsheetManager.initialize();
  
  // シートタブの切り替え処理
  sheetTabs.forEach((tab, index) => {
    tab.addEventListener('click', () => {
      sheetTabs.forEach(t => t.classList.remove('active'));
      tab.classList.add('active');
      KarteManager.changeSheet(index);
    });
  });
  
  // 保存ボタンのイベント
  saveButton.addEventListener('click', () => {
    KarteManager.saveKarte();
  });
  
  // 新規ボタンのイベント
  newButton.addEventListener('click', () => {
    KarteManager.createNewKarte();
  });
  
  // 開くボタンのイベント
  openButton.addEventListener('click', () => {
    KarteManager.loadKarteList();
    karteListModal.style.display = 'block';
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
    
    // input要素をリセット
    event.target.value = '';
  });
  
  // エクスポートボタンのイベント
  exportButton.addEventListener('click', () => {
    const filename = `カルテ_${new Date().toISOString().slice(0, 10)}.xlsx`;
    const success = SpreadsheetManager.exportToExcel(filename);
    
    if (!success) {
      alert('エクスポート中にエラーが発生しました');
    }
  });
  
  // 印刷ボタンのイベント
  printButton.addEventListener('click', () => {
    window.print();
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
});