/**
 * カルテ関連の機能を管理するモジュール
 */
const KarteManager = {
  // 現在のカルテID
  currentKarteId: null,
  
  // 現在のシートインデックス
  currentSheetIndex: 0,
  
  /**
   * カルテ一覧を読み込む
   */
  loadKarteList: function() {
    const karteListBody = document.getElementById('karte-list-body');
    karteListBody.innerHTML = '<tr><td colspan="8" class="text-center">読み込み中...</td></tr>';
    
    db.collection(APP_CONFIG.KARTE_COLLECTION).orderBy('createdAt', 'desc').get()
      .then(snapshot => {
        karteListBody.innerHTML = '';
        
        if (snapshot.empty) {
          karteListBody.innerHTML = '<tr><td colspan="8" class="text-center">カルテがありません</td></tr>';
          return;
        }
        
        snapshot.forEach(doc => {
          const data = doc.data();
          const karteInfo = data.karteInfo || this.extractKarteInfo(data.basicInfoSheet || []);
          
          const row = document.createElement('tr');
          row.innerHTML = `
            <td>${karteInfo.karteNo || '-'}</td>
            <td>${karteInfo.staffName || '-'}</td>
            <td>${karteInfo.customerName || '-'}</td>
            <td>${karteInfo.groupName || '-'}</td>
            <td>${karteInfo.stayDate || '-'}</td>
            <td>${karteInfo.destination || '-'}</td>
            <td>${karteInfo.totalPersons || '-'}</td>
            <td>
              <button class="action-btn edit-btn" data-id="${doc.id}">編集</button>
              <button class="action-btn delete-btn" data-id="${doc.id}">削除</button>
            </td>
          `;
          
          karteListBody.appendChild(row);
        });
        
        // 編集ボタンのイベント設定
        document.querySelectorAll('.edit-btn').forEach(button => {
          button.addEventListener('click', (e) => {
            const karteId = e.target.getAttribute('data-id');
            this.loadKarteData(karteId);
            document.getElementById('karte-list-modal').style.display = 'none';
          });
        });
        
        // 削除ボタンのイベント設定
        document.querySelectorAll('.delete-btn').forEach(button => {
          button.addEventListener('click', (e) => {
            const karteId = e.target.getAttribute('data-id');
            if (confirm('このカルテを削除してもよろしいですか？削除すると元に戻せません。')) {
              this.deleteKarte(karteId);
            }
          });
        });
      })
      .catch(error => {
        console.error('Error loading karte list:', error);
        karteListBody.innerHTML = `<tr><td colspan="8" class="text-center">エラーが発生しました: ${error.message}</td></tr>`;
      });
  },
  
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
          const karteInfo = data.karteInfo || this.extractKarteInfo(data.basicInfoSheet || []);
          
          // カルテIDの表示を更新
          document.getElementById('current-karte-id').textContent = `カルテNo: ${karteInfo.karteNo || karteId}`;
          
          // 現在のシートに対応するデータをロード
          this.changeSheet(this.currentSheetIndex);
          
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
   * カルテを削除する
   * @param {string} karteId - 削除するカルテID
   */
  deleteKarte: function(karteId) {
    db.collection(APP_CONFIG.KARTE_COLLECTION).doc(karteId).delete()
      .then(() => {
        alert('カルテを削除しました');
        this.loadKarteList(); // カルテ一覧を再読み込み
        
        // 現在編集中のカルテが削除された場合、新規作成モードにする
        if (this.currentKarteId === karteId) {
          this.currentKarteId = null;
          document.getElementById('current-karte-id').textContent = '新規カルテ';
          this.changeSheet(this.currentSheetIndex); // テンプレートをロード
          document.getElementById('last-saved').textContent = '保存されていません';
        }
      })
      .catch(error => {
        console.error('Error deleting karte:', error);
        alert(`削除エラー: ${error.message}`);
      });
  },
  
  /**
   * シートを変更する
   * @param {number} sheetIndex - シートのインデックス
   */
  changeSheet: function(sheetIndex) {
    this.currentSheetIndex = sheetIndex;
    
    // 現在のデータを保存（編集中のカルテがある場合）
    if (this.currentKarteId) {
      this.saveCurrentSheetData();
    }
    
    if (this.currentKarteId) {
      // カルテデータがある場合はそのデータをロード
      const sheetType = APP_CONFIG.SHEET_TYPES[sheetIndex];
      
      db.collection(APP_CONFIG.KARTE_COLLECTION).doc(this.currentKarteId).get()
        .then(doc => {
          if (doc.exists && doc.data()[sheetType]) {
            SpreadsheetManager.loadData(doc.data()[sheetType]);
          } else {
            // データがない場合はテンプレートをロード
            SpreadsheetManager.loadTemplateBySheetIndex(sheetIndex);
          }
        })
        .catch(error => {
          console.error(`Error loading ${sheetType}:`, error);
          // エラー時もテンプレートをロード
          SpreadsheetManager.loadTemplateBySheetIndex(sheetIndex);
        });
    } else {
      // 新規カルテの場合はテンプレートをロード
      SpreadsheetManager.loadTemplateBySheetIndex(sheetIndex);
    }
  },
  
  /**
   * 現在のシートデータを保存
   */
  saveCurrentSheetData: function() {
    if (!this.currentKarteId) return;
    
    const sheetData = SpreadsheetManager.getData();
    const sheetType = APP_CONFIG.SHEET_TYPES[this.currentSheetIndex];
    
    const updateData = {};
    updateData[sheetType] = sheetData;
    
    db.collection(APP_CONFIG.KARTE_COLLECTION).doc(this.currentKarteId).update(updateData)
      .then(() => {
        document.getElementById('last-saved').textContent = new Date().toLocaleString();
      })
      .catch(error => {
        console.error('Error saving sheet data:', error);
      });
  },
  
  /**
   * 新規カルテを作成
   */
  createNewKarte: function() {
    if (confirm('新規カルテを作成しますか？未保存の変更は失われます。')) {
      this.currentKarteId = null;
      document.getElementById('current-karte-id').textContent = '新規カルテ';
      
      // 現在のシートを初期化
      this.changeSheet(this.currentSheetIndex);
      document.getElementById('last-saved').textContent = '保存されていません';
    }
  },
  
  /**
   * カルテを保存する
   */
  saveKarte: function() {
    const sheetData = SpreadsheetManager.getData();
    
    // 基本情報シートからカルテ情報を抽出
    let karteInfo = {};
    if (this.currentSheetIndex === 0) {
      karteInfo = this.extractKarteInfo(sheetData);
    } else if (this.currentKarteId) {
      // 他のシートの場合は既存のカルテ情報を取得
      const getKarteInfo = async () => {
        try {
          const doc = await db.collection(APP_CONFIG.KARTE_COLLECTION).doc(this.currentKarteId).get();
          if (doc.exists) {
            return doc.data().karteInfo || this.extractKarteInfo(doc.data().basicInfoSheet || []);
          }
        } catch (error) {
          console.error('Error getting karte info:', error);
        }
        return {};
      };
      
      getKarteInfo().then(info => {
        karteInfo = info;
        this.continueKarteSave(sheetData, karteInfo);
      });
      return;
    }
    
    this.continueKarteSave(sheetData, karteInfo);
  },
  
  /**
   * カルテ保存の続き
   * @param {Array} sheetData - シートデータ
   * @param {Object} karteInfo - カルテ情報
   */
  continueKarteSave: function(sheetData, karteInfo) {
    if (!this.currentKarteId) {
      // 新規カルテの場合
      const karteNumber = prompt('新しいカルテ番号を入力してください:');
      if (!karteNumber) return;
      
      // 基本情報シートの場合、カルテ番号をセットする
      if (this.currentSheetIndex === 0) {
        SpreadsheetManager.setDataAtCell(1, 1, karteNumber);
        karteInfo.karteNo = karteNumber;
      } else {
        karteInfo.karteNo = karteNumber;
      }
      
      this.currentKarteId = db.collection(APP_CONFIG.KARTE_COLLECTION).doc().id;
    }
    
    document.getElementById('saving-indicator').style.display = 'block';
    
    // 現在のシートデータを準備
    const sheetType = APP_CONFIG.SHEET_TYPES[this.currentSheetIndex];
    const updateData = {
      karteInfo: karteInfo,
      [sheetType]: sheetData,
      lastUpdated: firebase.firestore.FieldValue.serverTimestamp()
    };
    
    // 新規カルテの場合は作成日時も設定
    if (!updateData.createdAt) {
      updateData.createdAt = firebase.firestore.FieldValue.serverTimestamp();
    }
    
    // Firestoreに保存
    db.collection(APP_CONFIG.KARTE_COLLECTION).doc(this.currentKarteId).set(updateData, { merge: true })
      .then(() => {
        document.getElementById('saving-indicator').style.display = 'none';
        document.getElementById('last-saved').textContent = new Date().toLocaleString();
        document.getElementById('current-karte-id').textContent = `カルテNo: ${karteInfo.karteNo || this.currentKarteId}`;
        alert('カルテを保存しました');
      })
      .catch(error => {
        document.getElementById('saving-indicator').style.display = 'none';
        alert(`保存エラー: ${error.message}`);
      });
  },
  
  /**
   * 基本情報シートからカルテ情報を抽出
   * @param {Array} sheetData - シートデータ
   * @return {Object} カルテ情報
   */
  extractKarteInfo: function(sheetData) {
    try {
      const karteInfo = {
        karteNo: sheetData[1][1] || '',
        staffName: sheetData[0][3] || '',
        customerName: sheetData[1][3] || '',
        groupName: sheetData[1][5] || '',
        stayDate: sheetData[2][1] || '',
        stayNights: sheetData[2][3] || '',
        totalPersons: sheetData[2][5] || '',
        destination: sheetData[3][3] || ''
      };
      return karteInfo;
    } catch (error) {
      console.error('Error extracting karte info:', error);
      return {};
    }
  }
};