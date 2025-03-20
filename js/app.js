/**
 * アプリケーションのメインエントリーポイント
 */
document.addEventListener('DOMContentLoaded', function() {
  console.log('DOMContentLoaded: アプリケーションを初期化します');
  
  // DOM要素
  const saveButton = document.getElementById('save-button');
  const newButton = document.getElementById('new-button');
  const openButton = document.getElementById('open-button');
  const exportButton = document.getElementById('export-button');
  const fileImport = document.getElementById('file-import');
  const closeModal = document.querySelector('.close-modal');
  const karteListModal = document.getElementById('karte-list-modal');
  
  // アプリケーションの状態変数
  let isInitialized = false;
  let lastChangeTime = 0;
  let autoSaveTimer = null;
  
  // カルテ検索機能の設定
  function setupKarteSearch() {
    const searchInput = document.getElementById('karte-search');
    const clearButton = document.getElementById('clear-search');
    
    if (!searchInput || !clearButton) return;
    
    // 検索実行
    searchInput.addEventListener('input', function() {
      const searchText = this.value.toLowerCase();
      filterKarteList(searchText);
    });
    
    // 検索クリア
    clearButton.addEventListener('click', function() {
      searchInput.value = '';
      filterKarteList('');
      searchInput.focus();
    });
    
    // 検索結果のフィルタリング
    function filterKarteList(searchText) {
      const rows = document.querySelectorAll('#karte-list-body tr');
      
      rows.forEach(row => {
        if (!searchText) {
          row.style.display = '';
          return;
        }
        
        const cells = row.querySelectorAll('td');
        let found = false;
        
        cells.forEach(cell => {
          if (cell.textContent.toLowerCase().includes(searchText)) {
            found = true;
          }
        });
        
        row.style.display = found ? '' : 'none';
      });
      
      // 検索結果がない場合のメッセージ
      const tbody = document.getElementById('karte-list-body');
      const noResults = document.getElementById('no-results-message');
      
      // 表示されている行があるか確認
      const hasVisibleRows = Array.from(rows).some(row => row.style.display !== 'none');
      
      // 検索結果がない場合のメッセージを表示
      if (!hasVisibleRows && searchText) {
        if (!noResults) {
          const message = document.createElement('tr');
          message.id = 'no-results-message';
          message.innerHTML = `<td colspan="9" style="text-align: center; padding: 20px;">検索結果が見つかりません。検索条件「${searchText}」に一致するカルテはありません。</td>`;
          tbody.appendChild(message);
        }
      } else if (noResults) {
        noResults.remove();
      }
    }
  }
  
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
      
      // 検索機能の設定
      setupKarteSearch();
      
      // 初期化後、カルテ一覧を自動的に表示
      setTimeout(() => {
        KarteManager.loadKarteList();
        karteListModal.style.display = 'block';
      }, 500);
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
  });
  
  // 出力ボタンのイベント
  exportButton.addEventListener('click', () => {
    try {
      console.log('カルテをエクスポートします');
      
      // カルテNoまたは現在の日時をファイル名に使用
      let filename = 'カルテ';
      
      // カルテNoが設定されている場合はそれを使用
      if (KarteManager.currentKarteId) {
        const karteNo = document.getElementById('current-karte-id').textContent;
        if (karteNo && karteNo !== '新規カルテ') {
          filename = karteNo.replace('カルテNo: ', '');
        }
      }
      
      // 現在の日時を追加
      const now = new Date();
      const dateStr = now.getFullYear() + 
                      ('0' + (now.getMonth() + 1)).slice(-2) + 
                      ('0' + now.getDate()).slice(-2) + '_' +
                      ('0' + now.getHours()).slice(-2) + 
                      ('0' + now.getMinutes()).slice(-2);
      
      // ファイル名を生成
      filename = `${filename}_${dateStr}.xlsx`;
      
      // エクスポート実行
      const result = SpreadsheetManager.exportToExcel(filename);
      
      if (result) {
        console.log('エクスポートが完了しました');
      } else {
        console.error('エクスポート中にエラーが発生しました');
        alert('エクスポート中にエラーが発生しました');
      }
    } catch (error) {
      console.error('エクスポート中にエラーが発生しました:', error);
      alert(`エクスポート中にエラーが発生しました: ${error.message}`);
    }
  });
  
  // ファイルインポート処理は残しておく（将来の拡張性のため）
  fileImport.addEventListener('change', (event) => {
    const file = event.target.files[0];
    if (!file) return;
    
    console.log(`ファイル "${file.name}" をインポートします`);
    
    SpreadsheetManager.importFromExcel(file, (success, errorMessage) => {
      if (success) {
        console.log('インポートが完了しました');
        
        // インポート後は新規カルテ状態に
        KarteManager.currentKarteId = null;
        document.getElementById('current-karte-id').textContent = '新規カルテ（インポート済み）';
        document.getElementById('last-saved').textContent = '保存されていません';
      } else {
        console.error('インポート中にエラーが発生しました:', errorMessage);
        alert(`インポート中にエラーが発生しました: ${errorMessage}`);
      }
      
      // インポート後にファイル選択をリセット
      fileImport.value = '';
    });
  });
  
  // モーダルを閉じるボタンのイベント
  closeModal.addEventListener('click', () => {
    karteListModal.style.display = 'none';
  });
  
  // モーダル外をクリックした時にも閉じる
  window.addEventListener('click', (event) => {
    if (event.target === karteListModal) {
      karteListModal.style.display = 'none';
    }
  });
  
  // ESCキーでモーダルを閉じる
  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape' && karteListModal.style.display === 'block') {
      karteListModal.style.display = 'none';
    }
  });
  
  // オフライン状態の検出
  window.addEventListener('offline', () => {
    document.getElementById('connection-status').textContent = 'オフライン';
    document.getElementById('connection-status').style.color = 'red';
    console.warn('ネットワーク接続がオフラインになりました');
  });
  
  // オンライン状態の検出
  window.addEventListener('online', () => {
    document.getElementById('connection-status').textContent = 'オンライン';
    document.getElementById('connection-status').style.color = '';
    console.log('ネットワーク接続が復旧しました');
  });
  
  // アプリケーション終了前の保存確認
  window.addEventListener('beforeunload', (event) => {
    // 変更があり、最後の変更から5秒以内の場合
    if (lastChangeTime > 0 && (Date.now() - lastChangeTime) < 5000) {
      const message = '変更が保存されていない可能性があります。ページを離れますか？';
      event.returnValue = message;
      return message;
    }
  });
});
