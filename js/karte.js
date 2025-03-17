/**
 * カルテデータを読み込む
 * @param {string} karteId - カルテID
 */
loadKarteData: function(karteId) {
  this.currentKarteId = karteId;
  
  db.collection(APP_CONFIG.KARTE_COLLECTION).doc(karteId).get()
    .then(doc => {
      if (doc.exists) {
        const data = doc.data();
        
        // シートデータを取得（シリアライズされたデータがあれば復元）
        const sheetData = this.parseSheetData(data);
        
        const karteInfo = data.karteInfo || this.extractKarteInfo(sheetData);
        
        // カルテIDの表示を更新
        document.getElementById('current-karte-id').textContent = `カルテNo: ${karteInfo.karteNo || karteId}`;
        
        // シートデータをロード
        SpreadsheetManager.loadData(sheetData);
        
        // カルテNoをシート内のセルにも設定
        if (karteInfo.karteNo) {
          // ロード後少し遅延させて設定（データロードが完了してから）
          setTimeout(() => {
            // カルテNo欄（通常3行目、B列）に値を設定
            SpreadsheetManager.setDataAtCell(2, 1, karteInfo.karteNo);
          }, 300);
        }
        
        // 最終更新日時の表示を更新
        if (data.lastUpdated) {
          const lastUpdated = data.lastUpdated.toDate();
          document.getElementById('last-saved').textContent = lastUpdated.toLocaleString();
        }
      } else {
        alert('カルテが見つかりません');
      }
    })
    .catch(error => {
      console.error('Error loading karte data:', error);
      alert(`読み込みエラー: ${error.message}`);
    });
}
