/**
 * Firebase設定
 * 実際のプロジェクト情報に置き換えてください
 */
const firebaseConfig = {
  apiKey: "AIzaSyBJmvoumJa1zcNBMNcJt4LP2Zjd2LbzEG0",
  authDomain: "travelkarute.firebaseapp.com",
  projectId: "travelkarute",
  storageBucket: "travelkarute.firebasestorage.app",
  messagingSenderId: "635350270309",
  appId: "1:635350270309:web:2498d21f9f134defeda18c",
  measurementId: "G-REHJZHXSR9"
};

// Firebaseの初期化
firebase.initializeApp(firebaseConfig);
const db = firebase.firestore();

/**
 * アプリケーション設定
 */
const APP_CONFIG = {
  // Firestoreのコレクション名
  KARTE_COLLECTION: 'karte',
  
  // カルテデータのフィールド名
  SHEET_DATA_FIELD: 'sheetData',
  
  // Handsontableの設定
  HOT_OPTIONS: {
    rowHeaders: true,
    colHeaders: true, // 列ヘッダー（A,B,C...）を表示するように変更
    contextMenu: true,
    manualColumnResize: true,
    manualRowResize: true,
    fixedRowsTop: 0,
    fixedColumnsLeft: 0,
    licenseKey: 'non-commercial-and-evaluation'
  }
};
