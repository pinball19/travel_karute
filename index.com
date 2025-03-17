<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Excel風カルテシステム</title>
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/handsontable/13.1.0/handsontable.min.css">
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
  <style>
    body {
      font-family: 'Segoe UI', Arial, sans-serif;
      margin: 0;
      padding: 0;
      background-color: #f5f5f5;
    }
    
    .excel-app {
      display: flex;
      flex-direction: column;
      height: 100vh;
    }
    
    .excel-header {
      background-color: #217346; /* Excel緑 */
      color: white;
      padding: 10px 20px;
      display: flex;
      justify-content: space-between;
      align-items: center;
    }
    
    .excel-ribbon {
      background-color: #f3f2f1;
      border-bottom: 1px solid #e1dfdd;
      padding: 5px 10px;
      display: flex;
      gap: 10px;
    }
    
    .ribbon-button {
      padding: 5px 10px;
      border: 1px solid transparent;
      background: none;
      border-radius: 3px;
      cursor: pointer;
      display: flex;
      flex-direction: column;
      align-items: center;
      font-size: 12px;
    }
    
    .ribbon-button:hover {
      border: 1px solid #c8c6c4;
      background-color: #f9f8f7;
    }
    
    .ribbon-button i {
      font-size: 16px;
      margin-bottom: 3px;
    }
    
    .ribbon-divider {
      width: 1px;
      background-color: #e1dfdd;
      margin: 0 5px;
    }
    
    .sheet-tabs {
      display: flex;
      background-color: #f3f2f1;
      border-top: 1px solid #e1dfdd;
      padding: 0 10px;
    }
    
    .sheet-tab {
      padding: 5px 15px;
      border: 1px solid #e1dfdd;
      border-bottom: none;
      border-radius: 3px 3px 0 0;
      margin-right: 2px;
      background-color: #f9f8f7;
      cursor: pointer;
      font-size: 13px;
    }
    
    .sheet-tab.active {
      background-color: white;
      border-top: 2px solid #217346;
    }
    
    .spreadsheet-container {
      flex-grow: 1;
      overflow: hidden;
    }
    
    .hot-container {
      width: 100%;
      height: 100%;
    }
    
    .status-bar {
      background-color: #f3f2f1;
      border-top: 1px solid #e1dfdd;
      padding: 5px 10px;
      font-size: 12px;
      display: flex;
      justify-content: space-between;
    }
    
    /* Handsontableをよりエクセルに近づけるためのスタイル */
    .excel-theme .handsontable {
      font-family: 'Segoe UI', Arial, sans-serif;
      font-size: 13px;
    }
    
    .excel-theme .handsontable .htCore td {
      border-color: #e0e0e0;
    }
    
    .excel-theme .handsontable .htCore th {
      background-color: #f3f2f1;
      color: #333;
      font-weight: normal;
    }
    
    .excel-theme .handsontable .htCore tbody tr:first-child td {
      border-top: 1px solid #e0e0e0;
    }
    
    .excel-theme .handsontable .htCore thead th:first-child {
      border-left: 1px solid #e0e0e0;
    }
    
    /* Firebase認証関連のスタイル */
    .auth-container {
      position: fixed;
      top: 0;
      left: 0;
      right: 0;
      bottom: 0;
      background-color: rgba(0, 0, 0, 0.5);
      display: flex;
      justify-content: center;
      align-items: center;
      z-index: 1000;
    }
    
    .auth-panel {
      background-color: white;
      padding: 30px;
      border-radius: 8px;
      width: 400px;
      box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
    }
    
    .auth-header {
      text-align: center;
      margin-bottom: 20px;
    }
    
    .auth-form {
      display: flex;
      flex-direction: column;
      gap: 15px;
    }
    
    .auth-input {
      padding: 10px;
      border: 1px solid #ddd;
      border-radius: 4px;
    }
    
    .auth-button {
      padding: 10px;
      background-color: #217346;
      color: white;
      border: none;
      border-radius: 4px;
      cursor: pointer;
    }
    
    .auth-button:hover {
      background-color: #1e6b3e;
    }
    
    .auth-alternative {
      text-align: center;
      margin-top: 15px;
      font-size: 14px;
    }
    
    .auth-link {
      color: #217346;
      cursor: pointer;
    }
  </style>
</head>
<body>
  <!-- Firebase認証パネル (初期状態では表示) -->
  <div id="auth-container" class="auth-container">
    <div class="auth-panel">
      <div class="auth-header">
        <h2>カルテシステム</h2>
        <p>続行するにはログインしてください</p>
      </div>
      <div id="login-form" class="auth-form">
        <input type="email" id="email-input" class="auth-input" placeholder="メールアドレス">
        <input type="password" id="password-input" class="auth-input" placeholder="パスワード">
        <button id="login-button" class="auth-button">ログイン</button>
        <div class="auth-alternative">
          アカウントをお持ちでない場合は<span id="show-signup" class="auth-link">登録</span>してください
        </div>
      </div>
      <div id="signup-form" class="auth-form" style="display: none;">
        <input type="email" id="signup-email" class="auth-input" placeholder="メールアドレス">
        <input type="password" id="signup-password" class="auth-input" placeholder="パスワード">
        <input type="password" id="signup-confirm" class="auth-input" placeholder="パスワード (確認)">
        <button id="signup-button" class="auth-button">アカウント作成</button>
        <div class="auth-alternative">
          すでにアカウントをお持ちの場合は<span id="show-login" class="auth-link">ログイン</span>してください
        </div>
      </div>
    </div>
  </div>
  
  <!-- Excel風アプリケーション（認証後に表示） -->
  <div class="excel-app" style="display: none;">
    <!-- ヘッダー部分 (Excelタイトルバー) -->
    <div class="excel-header">
      <div class="app-title">カルテシステム - [団体ナビ成約カルテ.xlsx]</div>
      <div class="user-info">
        <span id="user-email"></span>
        <button id="logout-button" style="margin-left: 10px;">ログアウト</button>
      </div>
    </div>
    
    <!-- リボン (Excel上部ツールバー) -->
    <div class="excel-ribbon">
      <button class="ribbon-button" id="save-button">
        <i class="fas fa-save"></i>
        保存
      </button>
      <button class="ribbon-button" id="new-button">
        <i class="fas fa-file"></i>
        新規
      </button>
      <div class="ribbon-divider"></div>
      <button class="ribbon-button" id="import-button">
        <i class="fas fa-file-import"></i>
        インポート
      </button>
      <button class="ribbon-button" id="export-button">
        <i class="fas fa-file-export"></i>
        エクスポート
      </button>
      <div class="ribbon-divider"></div>
      <button class="ribbon-button" id="print-button">
        <i class="fas fa-print"></i>
        印刷
      </button>
    </div>
    
    <!-- シートタブ (Excel下部のタブ) -->
    <div class="sheet-tabs">
      <div class="sheet-tab active">基本情報</div>
      <div class="sheet-tab">入金一覧</div>
      <div class="sheet-tab">支払一覧</div>
      <div class="sheet-tab">収支情報</div>
    </div>
    
    <!-- スプレッドシート本体 -->
    <div class="spreadsheet-container">
      <div id="hot-container" class="hot-container excel-theme"></div>
    </div>
    
    <!-- ステータスバー -->
    <div class="status-bar">
      <div>シート1</div>
      <div>最終更新: <span id="last-saved">保存されていません</span></div>
      <div>接続状態: <span id="connection-status">オンライン</span></div>
    </div>
  </div>
  
  <!-- 保存中インジケーター -->
  <div id="saving-indicator" style="display: none; position: fixed; bottom: 20px; right: 20px; background-color: #333; color: white; padding: 10px 20px; border-radius: 5px; box-shadow: 0 2px 10px rgba(0,0,0,0.2);">
    <i class="fas fa-spinner fa-spin"></i> 保存中...
  </div>
  
  <!-- 非表示のファイルインポート用input -->
  <input type="file" id="file-import" style="display: none;" accept=".xlsx, .xls, .csv">
  
  <!-- スクリプト -->
  <script src="https://cdnjs.cloudflare.com/ajax/libs/handsontable/13.1.0/handsontable.full.min.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/firebase/9.22.2/firebase-app-compat.min.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/firebase/9.22.2/firebase-firestore-compat.min.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/firebase/9.22.2/firebase-auth-compat.min.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
  
  <script>
    // Firebase設定 (実際のプロジェクト情報に置き換えてください)
    const firebaseConfig = {
      apiKey: "YOUR_API_KEY",
      authDomain: "your-project.firebaseapp.com",
      projectId: "your-project",
      storageBucket: "your-project.appspot.com",
      messagingSenderId: "YOUR_SENDER_ID",
      appId: "YOUR_APP_ID"
    };
    
    // Firebaseの初期化
    firebase.initializeApp(firebaseConfig);
    const db = firebase.firestore();
    const auth = firebase.auth();
    
    // DOMエレメント取得
    const authContainer = document.getElementById('auth-container');
    const loginForm = document.getElementById('login-form');
    const signupForm = document.getElementById('signup-form');
    const showSignupLink = document.getElementById('show-signup');
    const showLoginLink = document.getElementById('show-login');
    const loginButton = document.getElementById('login-button');
    const signupButton = document.getElementById('signup-button');
    const emailInput = document.getElementById('email-input');
    const passwordInput = document.getElementById('password-input');
    const signupEmail = document.getElementById('signup-email');
    const signupPassword = document.getElementById('signup-password');
    const signupConfirm = document.getElementById('signup-confirm');
    const userEmail = document.getElementById('user-email');
    const logoutButton = document.getElementById('logout-button');
    const excelApp = document.querySelector('.excel-app');
    const saveButton = document.getElementById('save-button');
    const newButton = document.getElementById('new-button');
    const importButton = document.getElementById('import-button');
    const exportButton = document.getElementById('export-button');
    const printButton = document.getElementById('print-button');
    const fileImport = document.getElementById('file-import');
    const sheetTabs = document.querySelectorAll('.sheet-tab');
    const savingIndicator = document.getElementById('saving-indicator');
    const lastSavedSpan = document.getElementById('last-saved');
    const connectionStatus = document.getElementById('connection-status');
    
    // 現在のカルテID
    let currentKarteId = null;
    
    // 現在のシートインデックス
    let currentSheetIndex = 0;
    
    // Handsontable インスタンス
    let hot = null;
    
    // 認証フォームの切り替え
    showSignupLink.addEventListener('click', () => {
      loginForm.style.display = 'none';
      signupForm.style.display = 'flex';
    });
    
    showLoginLink.addEventListener('click', () => {
      signupForm.style.display = 'none';
      loginForm.style.display = 'flex';
    });
    
    // ログイン処理
    loginButton.addEventListener('click', () => {
      const email = emailInput.value;
      const password = passwordInput.value;
      
      auth.signInWithEmailAndPassword(email, password)
        .then(() => {
          // ログイン成功
          showExcelApp();
        })
        .catch(error => {
          // ログインエラー
          alert(`ログインエラー: ${error.message}`);
        });
    });
    
    // サインアップ処理
    signupButton.addEventListener('click', () => {
      const email = signupEmail.value;
      const password = signupPassword.value;
      const confirm = signupConfirm.value;
      
      if (password !== confirm) {
        alert('パスワードが一致しません');
        return;
      }
      
      auth.createUserWithEmailAndPassword(email, password)
        .then(() => {
          // サインアップ成功
          showExcelApp();
        })
        .catch(error => {
          // サインアップエラー
          alert(`アカウント作成エラー: ${error.message}`);
        });
    });
    
    // ログアウト処理
    logoutButton.addEventListener('click', () => {
      auth.signOut()
        .then(() => {
          // ログアウト成功
          hideExcelApp();
        })
        .catch(error => {
          // ログアウトエラー
          alert(`ログアウトエラー: ${error.message}`);
        });
    });
    
    // 認証状態の監視
    auth.onAuthStateChanged(user => {
      if (user) {
        // ユーザーがログインしている場合
        userEmail.textContent = user.email;
        showExcelApp();
        loadKarteList();
      } else {
        // ユーザーがログアウトしている場合
        hideExcelApp();
      }
    });
    
    // Excel風アプリの表示
    function showExcelApp() {
      authContainer.style.display = 'none';
      excelApp.style.display = 'flex';
      
      // Handsontableが初期化されていない場合は初期化
      if (!hot) {
        initializeHandsontable();
      }
    }
    
    // Excel風アプリの非表示
    function hideExcelApp() {
      authContainer.style.display = 'flex';
      excelApp.style.display = 'none';
      emailInput.value = '';
      passwordInput.value = '';
      loginForm.style.display = 'flex';
      signupForm.style.display = 'none';
    }
    
    // Handsontableの初期化
    function initializeHandsontable() {
      const container = document.getElementById('hot-container');
      
      // 基本情報シートのテンプレート（初期状態）
      const basicInfoTemplate = [
        ['【団体ナビ成約カルテ】', '', '担当；', '', '記入日;', '', '個人通N0;', '', '変更その他報告事項'],
        ['カルテNo', '', '名前', '', '団体名', '', '電話', ''],
        ['宿泊日', '', '泊数', '', '合計\n人数', '', '成約日', ''],
        ['出発地', '', '⇒', '行先', '', '', '', ''],
        ['入金予定日', '入金場所', '入金予定額', '入金日', '入金額', '備考', 'チェック'],
        ['', '', '', '', '', '', ''],
        ['', '', '', '', '', '', ''],
        ['', '', '', '', '', '', ''],
        ['A；入金合計', '', '', '', '', '', ''],
        ['利用日', '手配先名；該当に〇を付ける', '電話/FAX', '担当者', '支払予定日', '支払金額', 'チェック'],
        ['', '', '', '', '', '', ''],
        ['', '', '', '', '', '', ''],
        ['', '', '', '', '', '', ''],
        ['B；支払合計', '', '', '', '', '', ''],
        ['報告日', '利益率', '利益額', '一人粗利', '旅行総額/A', '支払総額/B', '人数'],
        ['', '', '', '', '', '', ''],
      ];
      
      // Handsontableの設定
      hot = new Handsontable(container, {
        data: basicInfoTemplate,
        rowHeaders: true,
        colHeaders: true,
        contextMenu: true,
        manualColumnResize: true,
        manualRowResize: true,
        fixedRowsTop: 0,
        fixedColumnsLeft: 0,
        mergeCells: [
          {row: 0, col: 0, rowspan: 1, colspan: 2},
          {row: 8, col: 0, rowspan: 1, colspan: 7},
          {row: 13, col: 0, rowspan: 1, colspan: 7},
        ],
        cells: function(row, col) {
          const cellProperties = {};
          
          // ヘッダー行のスタイル
          if (row === 0 || row === 1 || row === 4 || row === 9 || row === 14) {
            cellProperties.renderer = headerRenderer;
          }
          
          // 合計行のスタイル
          if (row === 8 || row === 13) {
            cellProperties.renderer = totalRenderer;
          }
          
          return cellProperties;
        },
        licenseKey: 'non-commercial-and-evaluation'
      });
      
      // カスタムセルレンダラー：ヘッダーセル
      function headerRenderer(instance, td, row, col, prop, value, cellProperties) {
        Handsontable.renderers.TextRenderer.apply(this, arguments);
        td.style.backgroundColor = '#f3f2f1';
        td.style.fontWeight = 'bold';
        td.style.textAlign = 'center';
      }
      
      // カスタムセルレンダラー：合計セル
      function totalRenderer(instance, td, row, col, prop, value, cellProperties) {
        Handsontable.renderers.TextRenderer.apply(this, arguments);
        td.style.backgroundColor = '#e6f2ff';
        td.style.fontWeight = 'bold';
      }
      
      // ウィンドウのリサイズに合わせてHandsontableをリサイズ
      window.addEventListener('resize', () => {
        hot.updateSettings({
          width: container.offsetWidth,
          height: container.offsetHeight
        });
      });
    }
    
    // シートタブの切り替え
    sheetTabs.forEach((tab, index) => {
      tab.addEventListener('click', () => {
        sheetTabs.forEach(t => t.classList.remove('active'));
        tab.classList.add('active');
        currentSheetIndex = index;
        changeSheet(index);
      });
    });
    
    // シートの切り替え関数
    function changeSheet(index) {
      if (!hot) return;
      
      // 現在のデータを保存
      saveCurrentSheetData();
      
      // 選択されたシートによってデータを切り替え
      switch (index) {
        case 0: // 基本情報
          loadBasicInfoSheet();
          break;
        case 1: // 入金一覧
          loadPaymentSheet();
          break;
        case 2: // 支払一覧
          loadExpenseSheet();
          break;
        case 3: // 収支情報
          loadProfitSheet();
          break;
      }
    }
    
    // 基本情報シートのロード
    function loadBasicInfoSheet() {
      // カルテが選択されている場合はそのデータをロード
      if (currentKarteId) {
        db.collection('カルテ').doc(currentKarteId).get()
          .then(doc => {
            if (doc.exists && doc.data().basicInfoSheet) {
              hot.loadData(doc.data().basicInfoSheet);
            } else {
              // 基本的なテンプレートをロード
              const basicInfoTemplate = [
                ['【団体ナビ成約カルテ】', '', '担当；', '', '記入日;', '', '個人通N0;', '', '変更その他報告事項'],
                ['カルテNo', '', '名前', '', '団体名', '', '電話', ''],
                ['宿泊日', '', '泊数', '', '合計\n人数', '', '成約日', ''],
                ['出発地', '', '⇒', '行先', '', '', '', ''],
                ['入金予定日', '入金場所', '入金予定額', '入金日', '入金額', '備考', 'チェック'],
                ['', '', '', '', '', '', ''],
                ['', '', '', '', '', '', ''],
                ['', '', '', '', '', '', ''],
                ['A；入金合計', '', '', '', '', '', ''],
                ['利用日', '手配先名；該当に〇を付ける', '電話/FAX', '担当者', '支払予定日', '支払金額', 'チェック'],
                ['', '', '', '', '', '', ''],
                ['', '', '', '', '', '', ''],
                ['', '', '', '', '', '', ''],
                ['B；支払合計', '', '', '', '', '', ''],
                ['報告日', '利益率', '利益額', '一人粗利', '旅行総額/A', '支払総額/B', '人数'],
                ['', '', '', '', '', '', ''],
              ];
              hot.loadData(basicInfoTemplate);
            }
          })
          .catch(error => {
            console.error('Error loading basic info sheet:', error);
          });
      } else {
        // カルテが選択されていない場合は空のテンプレートをロード
        const basicInfoTemplate = [
          ['【団体ナビ成約カルテ】', '', '担当；', '', '記入日;', '', '個人通N0;', '', '変更その他報告事項'],
          ['カルテNo', '', '名前', '', '団体名', '', '電話', ''],
          ['宿泊日', '', '泊数', '', '合計\n人数', '', '成約日', ''],
          ['出発地', '', '⇒', '行先', '', '', '', ''],
          ['入金予定日', '入金場所', '入金予定額', '入金日', '入金額', '備考', 'チェック'],
          ['', '', '', '', '', '', ''],
          ['', '', '', '', '', '', ''],
          ['', '', '', '', '', '', ''],
          ['A；入金合計', '', '', '', '', '', ''],
          ['利用日', '手配先名；該当に〇を付ける', '電話/FAX', '担当者', '支払予定日', '支払金額', 'チェック'],
          ['', '', '', '', '', '', ''],
          ['', '', '', '', '', '', ''],
          ['', '', '', '', '', '', ''],
          ['B；支払合計', '', '', '', '', '', ''],
          ['報告日', '利益率', '利益額', '一人粗利', '旅行総額/A', '支払総額/B', '人数'],
          ['', '', '', '', '', '', ''],
        ];
        hot.loadData(basicInfoTemplate);
      }
    }
    
    // 入金一覧シートのロード
    function loadPaymentSheet() {
      if (currentKarteId) {
        db.collection('カルテ').doc(currentKarteId).get()
          .then(doc => {
            if (doc.exists && doc.data().paymentSheet) {
              hot.loadData(doc.data().paymentSheet);
            } else {
              // 入金一覧のテンプレート
              const paymentTemplate = [
                ['入金一覧', '', '', '', '', '', ''],
                ['入金予定日', '入金場所', '入金予定額', '入金日', '入金額', '備考', 'チェック'],
                ['', '', '', '', '', '', ''],
                ['', '', '', '', '', '', ''],
                ['', '', '', '', '', '', ''],
                ['', '', '', '', '', '', ''],
                ['合計', '', '0', '', '0', '', ''],
              ];
              hot.loadData(paymentTemplate);
            }
          })
          .catch(error => {
            console.error('Error loading payment sheet:', error);
          });
      } else {
        // テンプレート
        const paymentTemplate = [
          ['入金一覧', '', '', '', '', '', ''],
          ['入金予定日', '入金場所', '入金予定額', '入金日', '入金額', '備考', 'チェック'],
          ['', '', '', '', '', '', ''],
          ['', '', '', '', '', '', ''],
          ['', '', '', '', '', '', ''],
          ['', '', '', '', '', '', ''],
          ['合計', '', '0', '', '0', '', ''],
        ];
        hot.loadData(paymentTemplate);
      }
    }
    
    // 支払一覧シートのロード
    function loadExpenseSheet() {
      if (currentKarteId) {
        db.collection('カルテ').doc(currentKarteId).get()
          .then(doc => {
            if (doc.exists && doc.data().expenseSheet) {
              hot.loadData(doc.data().expenseSheet);
            } else {
              // 支払一覧のテンプレート
              const expenseTemplate = [
                ['支払一覧', '', '', '', '', '', ''],
                ['利用日', '手配先名', '電話/FAX', '担当者', '支払方法', '支払金額', '手配状況'],
                ['', '', '', '', '', '', ''],
                ['', '', '', '', '', '', ''],
                ['', '', '', '', '', '', ''],
                ['', '', '', '', '', '', ''],
                ['合計', '', '', '', '', '0', ''],
              ];
              hot.loadData(expenseTemplate);
            }
          })
          .catch(error => {
            console.error('Error loading expense sheet:', error);
          });
      } else {
        // テンプレート
        const expenseTemplate = [
          ['支払一覧', '', '', '', '', '', ''],
          ['利用日', '手配先名', '電話/FAX', '担当者', '支払方法', '支払金額', '手配状況'],
          ['', '', '', '', '', '', ''],
          ['', '', '', '', '', '', ''],
          ['', '', '', '', '', '', ''],
          ['', '', '', '', '', '', ''],
          ['合計', '', '', '', '', '0', ''],
        ];
        hot.loadData(expenseTemplate);
      }
    }
    
    // 収支情報シートのロード
    function loadProfitSheet() {
      if (currentKarteId) {
        db.collection('カルテ').doc(currentKarteId).get()
          .then(doc => {
            if (doc.exists && doc.data().profitSheet) {
              hot.loadData(doc.data().profitSheet);
            } else {
              // 収支情報のテンプレート
              const profitTemplate = [
                ['収支情報', '', '', '', '', '', ''],
                ['報告日', '利益率', '利益額', '一人粗利', '旅行総額/A', '支払総額/B', '人数'],
                ['', '', '', '', '', '', ''],
                ['', '', '', '', '', '', ''],
              ];
              hot.loadData(profitTemplate);
            }
          })
          .catch(error => {
            console.error('Error loading profit sheet:', error);
          });
      } else {
        // テンプレート
        const profitTemplate = [
          ['収支情報', '', '', '', '', '', ''],
          ['報告日', '利益率', '利益額', '一人粗利', '旅行総額/A', '支払総額/B', '人数'],
          ['', '', '', '', '', '', ''],
          ['', '', '', '', '', '', ''],
        ];
        hot.loadData(profitTemplate);
      }
    }
    
    // 現在のシートデータを保存
    function saveCurrentSheetData() {
      if (!currentKarteId || !hot) return;
      
      const sheetData = hot.getData();
      const sheetType = ['basicInfoSheet', 'paymentSheet', 'expenseSheet', 'profitSheet'][currentSheetIndex];
      
      const updateData = {};
      updateData[sheetType] = sheetData;
      
      db.collection('カルテ').doc(currentKarteId).update(updateData)
        .then(() => {
          lastSavedSpan.textContent = new Date().toLocaleString();
        })
        .catch(error => {
          console.error('Error saving sheet data:', error);
        });
    }
    
    // カルテ一覧を読み込む
    function loadKarteList() {
      // 実際の実装ではFirestoreからカルテ一覧を取得して
      // UIに表示する処理を実装します
    }
    
    // 保存ボタンのイベント
    saveButton.addEventListener('click', () => {
      if (!currentKarteId) {
        // 新規カルテの場合
        const karteNumber = prompt('新しいカルテ番号を入力してください:');
        if (!karteNumber) return;
        
        currentKarteId = db.collection('カルテ').doc().id;
        
        // 基本情報シートの場合、カルテ番号をセットする
        if (currentSheetIndex === 0) {
          hot.setDataAtCell(1, 1, karteNumber);
        }
      }
      
      savingIndicator.style.display = 'block';
      
      // 現在のシートデータを保存
      saveCurrentSheetData();
      
      // 他の必要なメタデータを保存
      db.collection('カルテ').doc(currentKarteId).set({
        lastUpdated: firebase.firestore.FieldValue.serverTimestamp(),
        updatedBy: auth.currentUser.email,
      }, { merge: true })
        .then(() => {
          savingIndicator.style.display = 'none';
          lastSavedSpan.textContent = new Date().toLocaleString();
          alert('カルテを保存しました');
        })
        .catch(error => {
          savingIndicator.style.display = 'none';
          alert(`保存エラー: ${error.message}`);
        });
    });
    
    // 新規ボタンのイベント
    newButton.addEventListener('click', () => {
      if (confirm('新規カルテを作成しますか？未保存の変更は失われます。')) {
        currentKarteId = null;
        
        // 現在のシートを初期化
        changeSheet(currentSheetIndex);
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
      
      const reader = new FileReader();
      reader.onload = function(e) {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          
          // 最初のシートのデータを取得
          const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
          const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
          
          // データをHandsontableに読み込む
          hot.loadData(jsonData);
          
          // インポート後は未保存状態
          currentKarteId = null;
          lastSavedSpan.textContent = '保存されていません';
          
          alert('Excelファイルをインポートしました');
        } catch (error) {
          alert(`インポートエラー: ${error.message}`);
        }
      };
      reader.readAsArrayBuffer(file);
      
      // input要素をリセット
      event.target.value = '';
    });
    
    // エクスポートボタンのイベント
    exportButton.addEventListener('click', () => {
      try {
        // 現在のシートデータを取得
        const data = hot.getData();
        
        // ワークシートを作成
        const ws = XLSX.utils.aoa_to_sheet(data);
        
        // 新しいワークブックを作成してシートを追加
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'カルテデータ');
        
        // Excelファイルとして出力
        const filename = `カルテ_${new Date().toISOString().slice(0, 10)}.xlsx`;
        XLSX.writeFile(wb, filename);
      } catch (error) {
        alert(`エクスポートエラー: ${error.message}`);
      }
    });
    
    // 印刷ボタンのイベント
    printButton.addEventListener('click', () => {
      window.print();
    });
    
    // オンライン状態の監視
    window.addEventListener('online', () => {
      connectionStatus.textContent = 'オンライン';
      connectionStatus.style.color = 'green';
    });
    
    window.addEventListener('offline', () => {
      connectionStatus.textContent = 'オフライン';
      connectionStatus.style.color = 'red';
    });
  </script>
</body>
</html>
