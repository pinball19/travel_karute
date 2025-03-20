/**
 * カルテデータを管理するモジュール
 */
const KarteManager = {
  // 現在のカルテID
  currentKarteId: null,
  
  // リアルタイムリスナー
  realtimeListener: null,
  
  // 最後のローカル更新タイムスタンプ
  lastLocalUpdate: 0,
  
  // 変更中フラグ（リアルタイム更新のループ防止用）
  isUpdating: false,
  
  /**
   * カルテデータを読み込む
   * @param {string} karteId - カルテID
   */
  loadKarteData: function(karteId) {
    this.currentKarteId = karteId;
    
    // 既存のリアルタイムリスナーがあれば削除
    this.detachRealtimeListener();
    
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
          
          // リアルタイムリスナーを設定
          this.attachRealtimeListener(karteId);
          
          // 編集者情報を表示
          if (data.currentEditors) {
            this.updateEditorsList(data.currentEditors);
          }
          
          // 現在の編集者として自分を追加
          this.registerCurrentEditor();
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
   * リアルタイム同期のリスナーを設定
   * @param {string} karteId - カルテID
   */
  attachRealtimeListener: function(karteId) {
    // 既存のリスナーがあれば削除
    this.detachRealtimeListener();
    
    // 新しいリスナーを設定
    this.realtimeListener = db.collection(APP_CONFIG.KARTE_COLLECTION).doc(karteId)
      .onSnapshot(docSnapshot => {
        // 自分自身の更新によるイベントは無視
        if (this.isUpdating) {
          return;
        }
        
        const data = docSnapshot.data();
        if (!data) return;
        
        // リモートの更新が最近のローカル更新より新しい場合のみ反映
        const remoteTimestamp = data.lastUpdated ? data.lastUpdated.toMillis() : 0;
        if (remoteTimestamp > this.lastLocalUpdate) {
          console.log('リモートから更新を受信しました');
          
          // シートデータを取得
          const sheetData = this.parseSheetData(data);
          
          // 現在の選択セルを記憶
          let selectedCell = null;
          if (SpreadsheetManager.hot) {
            selectedCell = SpreadsheetManager.hot.getSelected();
          }
          
          // データを更新
          SpreadsheetManager.loadData(sheetData);
          
          // 選択セルを復元
          if (selectedCell && SpreadsheetManager.hot) {
            setTimeout(() => {
              SpreadsheetManager.hot.selectCell(
                selectedCell[0][0],
                selectedCell[0][1],
                selectedCell[0][2],
                selectedCell[0][3]
              );
            }, 100);
          }
          
          // 最終更新日時の表示を更新
          if (data.lastUpdated) {
            const lastUpdated = data.lastUpdated.toDate();
            document.getElementById('last-saved').textContent = lastUpdated.toLocaleString();
            
            // 更新通知を表示
            this.showUpdateNotification();
          }
          
          // 編集者情報を更新
          if (data.currentEditors) {
            this.updateEditorsList(data.currentEditors);
          }
        }
      }, error => {
        console.error('リアルタイム同期エラー:', error);
      });
  },
  
  /**
   * リアルタイムリスナーを削除
   */
  detachRealtimeListener: function() {
    if (this.realtimeListener) {
      this.realtimeListener();
      this.realtimeListener = null;
    }
    
    // カルテを閉じる場合は編集者リストから自分を削除
    if (this.currentKarteId) {
      this.unregisterCurrentEditor();
    }
  },
  
  /**
   * 更新通知を表示
   */
  showUpdateNotification: function() {
    // 既存の通知があれば削除
    const existingNotification = document.getElementById('update-notification');
    if (existingNotification) {
      existingNotification.remove();
    }
    
    // 通知要素を作成
    const notification = document.createElement('div');
    notification.id = 'update-notification';
    notification.style.position = 'fixed';
    notification.style.bottom = '20px';
    notification.style.left = '20px';
    notification.style.backgroundColor = '#4CAF50';
    notification.style.color = 'white';
    notification.style.padding = '10px 20px';
    notification.style.borderRadius = '5px';
    notification.style.boxShadow = '0 2px 10px rgba(0,0,0,0.2)';
    notification.style.zIndex = '9999';
    notification.style.transition = 'opacity 0.5s';
    notification.innerHTML = '<i class="fas fa-sync-alt"></i> 他のユーザーによる変更が反映されました';
    
    // 通知を表示
    document.body.appendChild(notification);
    
    // 5秒後に通知を消す
    setTimeout(() => {
      notification.style.opacity = '0';
      setTimeout(() => {
        notification.remove();
      }, 500);
    }, 5000);
  },
  
  /**
   * 現在の編集者として登録
   */
  registerCurrentEditor: function() {
    if (!this.currentKarteId) return;
    
    // ユーザー識別子を取得（実際のアプリではログインユーザー情報を使用）
    const editorId = this.getEditorId();
    const editorName = this.getEditorName();
    
    // Firestoreの編集者リストに追加
    const karteRef = db.collection(APP_CONFIG.KARTE_COLLECTION).doc(this.currentKarteId);
    
    karteRef.get().then(doc => {
      if (doc.exists) {
        const data = doc.data();
        const currentEditors = data.currentEditors || {};
        
        // 自分の情報を追加
        currentEditors[editorId] = {
          name: editorName,
          lastActive: firebase.firestore.FieldValue.serverTimestamp()
        };
        
        // 5分以上更新のない編集者を削除
        const fiveMinutesAgo = new Date();
        fiveMinutesAgo.setMinutes(fiveMinutesAgo.getMinutes() - 5);
        
        Object.keys(currentEditors).forEach(id => {
          const editor = currentEditors[id];
          if (editor.lastActive && editor.lastActive.toDate() < fiveMinutesAgo) {
            delete currentEditors[id];
          }
        });
        
        // 更新
        karteRef.update({
          currentEditors: currentEditors
        }).catch(error => {
          console.error('編集者情報の更新エラー:', error);
        });
      }
    }).catch(error => {
      console.error('編集者情報の取得エラー:', error);
    });
    
    // 定期的に自分のアクティブ状態を更新
    this.startActiveStatusUpdater();
  },
  
  /**
   * 編集者リストから自分を削除
   */
  unregisterCurrentEditor: function() {
    if (!this.currentKarteId) return;
    
    // アクティブ状態の更新を停止
    this.stopActiveStatusUpdater();
    
    // ユーザー識別子を取得
    const editorId = this.getEditorId();
    
    // Firestoreの編集者リストから削除
    const karteRef = db.collection(APP_CONFIG.KARTE_COLLECTION).doc(this.currentKarteId);
    
    karteRef.get().then(doc => {
      if (doc.exists) {
        const data = doc.data();
        const currentEditors = data.currentEditors || {};
        
        // 自分の情報を削除
        if (currentEditors[editorId]) {
          delete currentEditors[editorId];
          
          // 更新
          karteRef.update({
            currentEditors: currentEditors
          }).catch(error => {
            console.error('編集者情報の更新エラー:', error);
          });
        }
      }
    }).catch(error => {
      console.error('編集者情報の取得エラー:', error);
    });
  },
  
  // アクティブ状態更新タイマー
  activeStatusTimer: null,
  
  /**
   * 定期的にアクティブ状態を更新
   */
  startActiveStatusUpdater: function() {
    // 既存のタイマーがあれば停止
    this.stopActiveStatusUpdater();
    
    // 1分ごとにアクティブ状態を更新
    this.activeStatusTimer = setInterval(() => {
      if (!this.currentKarteId) {
        this.stopActiveStatusUpdater();
        return;
      }
      
      const editorId = this.getEditorId();
      const karteRef = db.collection(APP_CONFIG.KARTE_COLLECTION).doc(this.currentKarteId);
      
      karteRef.get().then(doc => {
        if (doc.exists) {
          const data = doc.data();
          const currentEditors = data.currentEditors || {};
          
          if (currentEditors[editorId]) {
            // 最終アクティブ時間を更新
            currentEditors[editorId].lastActive = firebase.firestore.FieldValue.serverTimestamp();
            
            // 更新
            karteRef.update({
              currentEditors: currentEditors
            }).catch(error => {
              console.error('アクティブ状態の更新エラー:', error);
            });
          }
        }
      }).catch(error => {
        console.error('アクティブ状態の更新エラー:', error);
      });
    }, 60000); // 1分ごと
  },
  
  /**
   * アクティブ状態の更新を停止
   */
  stopActiveStatusUpdater: function() {
    if (this.activeStatusTimer) {
      clearInterval(this.activeStatusTimer);
      this.activeStatusTimer = null;
    }
  },
  
  /**
   * 編集者一覧を更新
   */
  updateEditorsList: function(editors) {
    // 編集者情報を表示するUI要素
    let editorsContainer = document.getElementById('current-editors');
    
    // まだ存在しない場合は作成
    if (!editorsContainer) {
      editorsContainer = document.createElement('div');
      editorsContainer.id = 'current-editors';
      editorsContainer.className = 'current-editors-container';
      
      // ヘッダー情報の隣に配置
      const infoElement = document.querySelector('.excel-header .info');
      if (infoElement) {
        infoElement.appendChild(editorsContainer);
      }
    }
    
    // 編集者リストを生成
    const editorsList = Object.keys(editors).map(id => {
      const editor = editors[id];
      return `<span class="editor-badge">${editor.name}</span>`;
    });
    
    // UI更新
    if (editorsList.length > 0) {
      editorsContainer.innerHTML = '編集中: ' + editorsList.join(' ');
      editorsContainer.style.display = 'block';
    } else {
      editorsContainer.style.display = 'none';
    }
  },
  
  /**
   * 編集者IDを取得
   * 実際のアプリではユーザー認証情報を使用する
   */
  getEditorId: function() {
    // ローカルストレージから取得、なければ新規生成
    let editorId = localStorage.getItem('editorId');
    if (!editorId) {
      editorId = 'editor_' + Math.random().toString(36).substr(2, 9);
      localStorage.setItem('editorId', editorId);
    }
    return editorId;
  },
  
  /**
   * 編集者名を取得
   * 実際のアプリではユーザー認証情報を使用する
   */
  getEditorName: function() {
    // ローカルストレージから取得、なければ新規生成
    let editorName = localStorage.getItem('editorName');
    if (!editorName) {
      // ランダムな名前を生成（実際のアプリではユーザー名を使用）
      const names = ['鈴木', '田中', '佐藤', '高橋', '渡辺'];
      editorName = names[Math.floor(Math.random() * names.length)] + ' ' + 
                   Math.floor(Math.random() * 100);
      localStorage.setItem('editorName', editorName);
    }
    return editorName;
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
    // 現在のカルテから編集者を削除
    this.detachRealtimeListener();
    
    // 現在のカルテIDをクリア
    this.currentKarteId = null;
    
    // 空のテンプレートをロード
    SpreadsheetManager.loadUnifiedTemplate();
    
    // カルテIDの表示を更新
    document.getElementById('current-karte-id').textContent = '新規カルテ';
    
    // 最終更新日時の表示をリセット
    document.getElementById('last-saved').textContent = '保存されていません';
    
    // 編集者リストをクリア
    const editorsContainer = document.getElementById('current-editors');
    if (editorsContainer) {
      editorsContainer.style.display = 'none';
    }
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
    
    // 現在の編集者情報を保持
    if (this.currentKarteId) {
      // 既存のカルテの場合、編集者情報を取得して保持
      db.collection(APP_CONFIG.KARTE_COLLECTION).doc(this.currentKarteId).get()
        .then(doc => {
          if (doc.exists) {
            const data = doc.data();
            if (data.currentEditors) {
              karteData.currentEditors = data.currentEditors;
            }
          }
          this.performSave(karteData);
        })
        .catch(error => {
          console.error('編集者情報の取得エラー:', error);
          this.performSave(karteData);
        });
    } else {
      // 新規カルテの場合
      const editorId = this.getEditorId();
      const editorName = this.getEditorName();
      
      karteData.currentEditors = {
        [editorId]: {
          name: editorName,
          lastActive: firebase.firestore.FieldValue.serverTimestamp()
        }
      };
      
      this.performSave(karteData);
    }
  },
  
  /**
   * 実際の保存処理を実行
   * @param {Object} karteData - 保存するカルテデータ
   */
  performSave: function(karteData) {
    // 更新中フラグをセット（自分の更新によるイベントを無視するため）
    this.isUpdating = true;
    
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
          
          // リアルタイムリスナーを設定
          this.attachRealtimeListener(docRef.id);
          
          return docRef;
        });
    }
    
    // 保存処理
    savePromise
      .then(() => {
        console.log('カルテを保存しました');
        
        // 最終ローカル更新タイムスタンプを更新
        this.lastLocalUpdate = Date.now();
        
        // カルテIDの表示を更新
        const displayId = karteData.karteInfo.karteNo || this.currentKarteId;
        document.getElementById('current-karte-id').textContent = `カルテNo: ${displayId}`;
        
        // 最終更新日時の表示を更新
        const now = new Date();
        document.getElementById('last-saved').textContent = now.toLocaleString();
        
        // 保存中インジケーターを非表示
        document.getElementById('saving-indicator').style.display = 'none';
        
        // 更新中フラグを解除
        setTimeout(() => {
          this.isUpdating = false;
        }, 1000);
      })
      .catch(error => {
        console.error('保存中にエラーが発生しました:', error);
        
        // 保存中インジケーターを非表示
        document.getElementById('saving-indicator').style.display = 'none';
        
        // 更新中フラグを解除
        this.isUpdating = false;
        
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
          karteListBody.innerHTML = '<tr><td colspan="9">カルテがありません</td></tr>';
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
          
          // 編集者
          const cellEditors = document.createElement('td');
          if (data.currentEditors && Object.keys(data.currentEditors).length > 0) {
            const editorNames = Object.values(data.currentEditors).map(editor => editor.name);
            cellEditors.textContent = editorNames.join(', ');
            cellEditors.style.color = '#e53935'; // 赤色で表示
          } else {
            cellEditors.textContent = '-';
          }
          row.appendChild(cellEditors);
          
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
        karteListBody.innerHTML = `<tr><td colspan="9">読み込みエラー: ${error.message}</td></tr>`;
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
  },
  
  /**
   * ページ離脱時の処理
   */
  handleBeforeUnload: function() {
    // 編集者リストから自分を削除
    this.unregisterCurrentEditor();
    
    // リアルタイムリスナーを削除
    this.detachRealtimeListener();
  }
};

// ページ離脱時の処理を設定
window.addEventListener('beforeunload', function() {
  KarteManager.handleBeforeUnload();
});
