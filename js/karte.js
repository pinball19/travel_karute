/**
 * カルテデータを管理するモジュール
 */
const KarteManager = {
  // 現在のカルテID
  currentKarteId: null,
  
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
  },
  
  /**
   * シートデータを解析する
   * @param {Object} data - カルテのデータオブジェクト
   * @return {Array} シートデータ
   */
  parseSheetData: function(data) {
    // シリアライズされたデータがない場合はテンプレートを返す
    if (!data || !data.serializedSheetData) {
      // 古い形式のデータ互換性のために確認
      if (data && data[APP_CONFIG.SHEET_DATA_FIELD]) {
        return data[APP_CONFIG.SHEET_DATA_FIELD];
      }
      return UNIFIED_TEMPLATE;
    }
    
    const serialized = data.serializedSheetData;
    const result = [];
    
    // 行数を特定
    const rowKeys = Object.keys(serialized)
      .filter(key => key.startsWith('row'))
      .sort((a, b) => {
        const numA = parseInt(a.substring(3));
        const numB = parseInt(b.substring(3));
        return numA - numB;
      });
    
    // 各行を復元
    rowKeys.forEach(rowKey => {
      const rowObj = serialized[rowKey];
      const row = [];
      
      // 列数を特定
      const colKeys = Object.keys(rowObj)
        .filter(key => key.startsWith('col'))
        .sort((a, b) => {
          const numA = parseInt(a.substring(3));
          const numB = parseInt(b.substring(3));
          return numA - numB;
        });
      
      // 各列を復元
      colKeys.forEach(colKey => {
        row.push(rowObj[colKey]);
      });
      
      result.push(row);
    });
    
    // 結果が空の場合はテンプレートを返す
    if (result.length === 0) {
      return UNIFIED_TEMPLATE;
    }
    
    return result;
  },
  
  /**
   * シートデータからカルテ情報を抽出する
   * @param {Array} sheetData - シートデータ
   * @return {Object} カルテ情報
   */
  extractKarteInfo: function(sheetData) {
    const info = {};
    
    try {
      // カルテNo (3行目、B列)
      if (sheetData[2] && sheetData[2][1]) {
        info.karteNo = sheetData[2][1];
      }
      
      // 担当者 (2行目、D列)
      if (sheetData[1] && sheetData[1][3]) {
        info.tantosha = sheetData[1][3];
      }
      
      // 名前 (3行目、D列)
      if (sheetData[2] && sheetData[2][3]) {
        info.name = sheetData[2][3];
      }
      
      // 団体名 (3行目、F列)
      if (sheetData[2] && sheetData[2][5]) {
        info.dantaiName = sheetData[2][5];
      }
      
      // 宿泊日 (4行目、B列)
      if (sheetData[3] && sheetData[3][1]) {
        info.stayDate = sheetData[3][1];
      }
      
      // 行先 (5行目、E列)
      if (sheetData[4] && sheetData[4][4]) {
        info.destination = sheetData[4][4];
      }
      
      // 人数 (4行目、F列)
      if (sheetData[3] && sheetData[3][5]) {
        info.personCount = sheetData[3][5];
      }
    } catch (error) {
      console.error('カルテ情報の抽出中にエラーが発生しました:', error);
    }
    
    return info;
  },
  
  /**
   * 新規カルテを作成する
   */
  createNewKarte: function() {
    // 現在のカルテIDをクリア
    this.currentKarteId = null;
    
    // 空のテンプレートをロード
    SpreadsheetManager.loadUnifiedTemplate();
    
    // カルテIDの表示を更新
    document.getElementById('current-karte-id').textContent = '新規カルテ';
    
    // 最終更新日時の表示をリセット
    document.getElementById('last-saved').textContent = '保存されていません';
  },
  
  /**
   * 現在のカルテを保存する
   */
  saveKarte: function() {
    // 保存中インジケーターを表示
    document.getElementById('saving-indicator').style.display = 'block';
    
    // 現在のシートデータを取得
    const sheetData = SpreadsheetManager.getData();
    
    // カルテ情報を抽出
    const karteInfo = this.extractKarteInfo(sheetData);
    
    // Firestoreに保存するためにシートデータを適切な形式に変換
    // 2次元配列を直接保存せず、シリアライズして保存する
    const serializedSheetData = this.serializeSheetData(sheetData);
    
    // 保存するデータを準備
    const karteData = {
      karteInfo: karteInfo,
      // シリアライズしたデータを保存
      serializedSheetData: serializedSheetData,
      lastUpdated: firebase.firestore.FieldValue.serverTimestamp()
    };
    
    // 保存先を決定（既存のカルテなら更新、新規なら作成）
    let savePromise;
    
    if (this.currentKarteId) {
      // 既存のカルテを更新
      savePromise = db.collection(APP_CONFIG.KARTE_COLLECTION)
        .doc(this.currentKarteId)
        .update(karteData);
    } else {
      // 新規カルテを作成
      savePromise = db.collection(APP_CONFIG.KARTE_COLLECTION)
        .add(karteData)
        .then(docRef => {
          this.currentKarteId = docRef.id;
          return docRef;
        });
    }
    
    // 保存処理
    savePromise
      .then(() => {
        console.log('カルテを保存しました');
        
        // カルテIDの表示を更新
        const displayId = karteInfo.karteNo || this.currentKarteId;
        document.getElementById('current-karte-id').textContent = `カルテNo: ${displayId}`;
        
        // 最終更新日時の表示を更新
        const now = new Date();
        document.getElementById('last-saved').textContent = now.toLocaleString();
        
        // 保存中インジケーターを非表示
        document.getElementById('saving-indicator').style.display = 'none';
      })
      .catch(error => {
        console.error('保存中にエラーが発生しました:', error);
        
        // 保存中インジケーターを非表示
        document.getElementById('saving-indicator').style.display = 'none';
        
        alert(`保存中にエラーが発生しました: ${error.message}`);
      });
  },
  
  /**
   * シートデータをFirestore用にシリアライズする
   * @param {Array} sheetData - 2次元配列のシートデータ
   * @return {Object} シリアライズしたデータ
   */
  serializeSheetData: function(sheetData) {
    // 2次元配列を各行を文字列に変換して保存
    const serialized = {};
    
    sheetData.forEach((row, rowIndex) => {
      // 各行をオブジェクトに変換
      const rowObj = {};
      row.forEach((cell, colIndex) => {
        // null や undefined の場合は空文字に変換
        rowObj[`col${colIndex}`] = cell != null ? String(cell) : '';
      });
      serialized[`row${rowIndex}`] = rowObj;
    });
    
    return serialized;
  },
  
  /**
   * カルテ一覧を読み込む
   */
  loadKarteList: function() {
    const karteListBody = document.getElementById('karte-list-body');
    
    // 一覧をクリア
    karteListBody.innerHTML = '';
    
    // カルテ一覧を取得
    db.collection(APP_CONFIG.KARTE_COLLECTION)
      .orderBy('lastUpdated', 'desc')
      .limit(50) // 最大50件まで表示
      .get()
      .then(querySnapshot => {
        if (querySnapshot.empty) {
          karteListBody.innerHTML = '<tr><td colspan="8">カルテがありません</td></tr>';
          return;
        }
        
        // 各カルテのデータを一覧に追加
        querySnapshot.forEach(doc => {
          const data = doc.data();
          const karteInfo = data.karteInfo || {};
          
          // 行要素を作成
          const row = document.createElement('tr');
          
          // カルテNo
          const cellKarteNo = document.createElement('td');
          cellKarteNo.textContent = karteInfo.karteNo || doc.id;
          row.appendChild(cellKarteNo);
          
          // 担当者
          const cellTantosha = document.createElement('td');
          cellTantosha.textContent = karteInfo.tantosha || '-';
          row.appendChild(cellTantosha);
          
          // 名前
          const cellName = document.createElement('td');
          cellName.textContent = karteInfo.name || '-';
          row.appendChild(cellName);
          
          // 団体名
          const cellDantai = document.createElement('td');
          cellDantai.textContent = karteInfo.dantaiName || '-';
          row.appendChild(cellDantai);
          
          // 宿泊日
          const cellStayDate = document.createElement('td');
          cellStayDate.textContent = karteInfo.stayDate || '-';
          row.appendChild(cellStayDate);
          
          // 行先
          const cellDestination = document.createElement('td');
          cellDestination.textContent = karteInfo.destination || '-';
          row.appendChild(cellDestination);
          
          // 人数
          const cellPersonCount = document.createElement('td');
          cellPersonCount.textContent = karteInfo.personCount || '-';
          row.appendChild(cellPersonCount);
          
          // 操作ボタン
          const cellActions = document.createElement('td');
          
          // 編集ボタン
          const editBtn = document.createElement('button');
          editBtn.className = 'action-btn edit-btn';
          editBtn.textContent = '編集';
          editBtn.onclick = () => {
            this.loadKarteData(doc.id);
            document.getElementById('karte-list-modal').style.display = 'none';
          };
          cellActions.appendChild(editBtn);
          
          // 削除ボタン
          const deleteBtn = document.createElement('button');
          deleteBtn.className = 'action-btn delete-btn';
          deleteBtn.textContent = '削除';
          deleteBtn.onclick = () => {
            if (confirm('このカルテを削除してもよろしいですか？')) {
              this.deleteKarte(doc.id);
            }
          };
          cellActions.appendChild(deleteBtn);
          
          row.appendChild(cellActions);
          
          // 行を一覧に追加
          karteListBody.appendChild(row);
        });
      })
      .catch(error => {
        console.error('カルテ一覧の読み込み中にエラーが発生しました:', error);
        karteListBody.innerHTML = `<tr><td colspan="8">読み込みエラー: ${error.message}</td></tr>`;
      });
  },
  
  /**
   * カルテを削除する
   * @param {string} karteId - 削除するカルテのID
   */
  deleteKarte: function(karteId) {
    db.collection(APP_CONFIG.KARTE_COLLECTION).doc(karteId).delete()
      .then(() => {
        console.log('カルテを削除しました');
        
        // 一覧を再読み込み
        this.loadKarteList();
        
        // 現在表示中のカルテが削除されたカルテの場合、新規カルテを作成
        if (this.currentKarteId === karteId) {
          this.createNewKarte();
        }
      })
      .catch(error => {
        console.error('削除中にエラーが発生しました:', error);
        alert(`削除中にエラーが発生しました: ${error.message}`);
      });
  }
};
