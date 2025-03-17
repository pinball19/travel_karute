/**
 * カルテ関連の機能を管理するモジュール
 */
const KarteManager = {
  // 現在のカルテID
  currentKarteId: null,
  
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
          const karteInfo = data.karteInfo || this.extractKarteInfo(this.parseSheetData(data));
          
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
   * Firestoreからのデータを解析
   * @param {Object} data - Firestoreから取得したデータ
   * @return {Array} シートデータの2次元配列
   */
  parseSheetData: function(data) {
    // シリアライズされたデータがある場合は復元
    if (data.sheetDataSerialized) {
      try {
        return JSON.parse(data.sheetDataSerialized);
      } catch (error) {
        console.error('Sheet data parsing error:', error);
        return UNIFIED_TEMPLATE; // エラー時はテンプレートを返す
      }
    }
    
    // 旧形式のデータがある場合
    if (data.sheetData) {
      return data.sheetData;
    }
    
    // 基本情報シートがある場合
    if (data.basicInfoSheet) {
      return data.basicInfoSheet;
    }
    
    // データがない場合はテンプレートを返す
    return UNIFIED_TEMPLATE;
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
          
          // シートデータを取得（シリアライズされたデータがあれば復元）
          const sheetData = this.parseSheetData(data);
          
          const karteInfo = data.karteInfo || this.extractKarteInfo(sheetData);
          
          // カルテIDの表示を更新
          document.getElementById('current-karte-id').textContent = `カルテNo: ${karteInfo.karteNo || karteId}`;
          
          // シートデータをロード
          SpreadsheetManager.loadData(sheetData);
          
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
          SpreadsheetManager.loadUnifiedTemplate(); // テンプレートをロード
          document.getElementById('last-saved').textContent = '保存されていません';
        }
      })
      .catch(error => {
        console.error('Error deleting karte:', error);
        alert(`削除エラー: ${error.message}`);
      });
  },
  
  /**
   * 新規カルテを作成
   */
  createNewKarte: function() {
    if (confirm('新規カルテを作成しますか？未保存の変更は失われます。')) {
      this.currentKarteId = null;
      document.getElementById('current-karte-id').textContent = '新規カルテ';
      
      // テンプレートをロード
      SpreadsheetManager.loadUnifiedTemplate();
      document.getElementById('last-saved').textContent = '保存されていません';
    }
  },
  
  /**
   * カルテを保存する
   */
  saveKarte: function() {
    try {
      // シートデータを取得
      const sheetData = SpreadsheetManager.getData();
      
      // カルテ情報を抽出
      const karteInfo = this.extractKarteInfo(sheetData);
      
      if (!this.currentKarteId) {
        // 新規カルテの場合
        const karteNumber = prompt('新しいカルテ番号を入力してください:');
        if (!karteNumber) return;
        
        // カルテ番号をセットする
        SpreadsheetManager.setDataAtCell(2, 1, karteNumber);
        karteInfo.karteNo = karteNumber;
        
        this.currentKarteId = db.collection(APP_CONFIG.KARTE_COLLECTION).doc().id;
      }
      
      document.getElementById('saving-indicator').style.display = 'block';
      
      // Firestoreに保存可能な形式に変換
      // 二次元配列をJSON文字列に変換
      const sheetDataSerialized = JSON.stringify(sheetData);
      
      // 保存データの準備
      const updateData = {
        karteInfo: karteInfo,
        sheetDataSerialized: sheetDataSerialized, // JSON文字列として保存
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
          console.error('保存エラー:', error);
          alert(`保存エラー: ${error.message}`);
        });
    } catch (error) {
      console.error('保存処理中にエラーが発生しました:', error);
      document.getElementById('saving-indicator').style.display = 'none';
      alert(`保存エラー: ${error.message}`);
    }
  },
  
  /**
   * シートデータからカルテ情報を抽出
   * @param {Array} sheetData - シートデータ
   * @return {Object} カルテ情報
   */
  extractKarteInfo: function(sheetData) {
    try {
      const karteInfo = {
        karteNo: sheetData[2][1] || '',        // カルテNo
        staffName: sheetData[1][3] || '',      // 担当
        customerName: sheetData[2][3] || '',   // 名前
        groupName: sheetData[2][5] || '',      // 団体名
        stayDate: sheetData[3][1] || '',       // 宿泊日
        stayNights: sheetData[3][3] || '',     // 泊数
        totalPersons: sheetData[3][5] || '',   // 合計人数
        destination: sheetData[4][3] || ''     // 行先
      };
      return karteInfo;
    } catch (error) {
      console.error('Error extracting karte info:', error);
      return {};
    }
  }
};
