/**
 * スプレッドシートのコメントを監視し、特定の相手へのメンションがあればメール通知する
 * ※ このスクリプトを動かすには、「Drive API (v3)」を追加が必要です。
 */
function checkCommentsAndNotify() {
  // --- 設定エリア ---
  const TARGET_EMAIL = 'j_kibushi@carecom.co.jp'; // テスト用：自分宛
  // 以下のキーワードがコメントに含まれている場合のみメールを送ります（誤送信防止）
  const TARGET_NAME_KEYWORDS = ['kibushi', '木伏', TARGET_EMAIL]; 
  const SHEET_NAME = '依頼一覧'; // 監視対象のシート名
  // ------------------

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fileId = ss.getId();
  
  // プロパティサービスから前回のチェック時間を取得
  const props = PropertiesService.getScriptProperties();
  const lastCheckTimeStr = props.getProperty('LAST_CHECK_TIME');
  
  // 現在時刻
  const now = new Date();

  // 初回実行時は、現在時刻を保存して終了（過去のコメントを大量に通知しないため）
  if (!lastCheckTimeStr) {
    props.setProperty('LAST_CHECK_TIME', now.toISOString());
    console.log("初回実行：基準時刻を保存しました。次回実行時から通知判定を行います。");
    return;
  }

  const lastCheckTime = new Date(lastCheckTimeStr);

  // Drive APIを使ってコメントを取得
  // 注意: Drive API v3 が有効になっている必要があります
  let comments = [];
  let pageToken = null;
  
  try {
    do {
      // startModifiedTime: 前回チェック時以降に更新されたコメントを取得
      const response = Drive.Comments.list(fileId, {
        startModifiedTime: lastCheckTime.toISOString(),
        pageSize: 20,
        pageToken: pageToken,
        fields: 'nextPageToken, comments(content, createdTime, author(emailAddress), resolved, anchor)'
      });
      
      if (response.comments && response.comments.length > 0) {
        comments = comments.concat(response.comments);
      }
      pageToken = response.nextPageToken;
    } while (pageToken);
  } catch (e) {
    console.error("Drive API エラー: " + e.toString());
    console.error("左側の「サービス」から「Drive API (v2)」を追加してください。");
    return;
  }

  // シートを取得
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    console.warn(`シート「${SHEET_NAME}」が見つかりません。`);
    return;
  }

  // 取得したコメントをチェック
  comments.forEach(function(comment) {
    const commentDate = new Date(comment.createdTime);
    
    // 1. 時間チェック（APIのフィルタ漏れ対策）
    if (commentDate <= lastCheckTime) return;

    // 2. 自分のコメントかチェック（自分が書いたものだけ通知対象にする）
    if (comment.author.emailAddress !== Session.getActiveUser().getEmail()) return;

    // 3. ステータスチェック（解決済みでないもの）
    if (comment.resolved) return;

    // 4. キーワードチェック（相手へのメンションが含まれているか）
    const content = comment.content || "";
    const hasKeyword = TARGET_NAME_KEYWORDS.some(keyword => content.includes(keyword));
    
    // キーワードが含まれていない場合はスキップ
    if (!hasKeyword) return;

    // 5. 行番号の特定（アンカー情報の解析）
    let row = -1;
    if (comment.anchor) {
      try {
        // anchorは通常 JSON文字列: {"r":5, "c":2} (0-indexed)
        const anchorData = JSON.parse(comment.anchor);
        if (typeof anchorData.r !== 'undefined') {
          row = anchorData.r + 1; // 0-indexed なので +1
        }
      } catch (e) {
        console.warn("アンカー情報の解析に失敗: " + comment.anchor);
      }
    }

    // 行が特定できたらメール送信
    if (row > 0) {
      // B列（2列目）の値を取得
      const subjectValue = sheet.getRange(row, 2).getValue();
      
      // メール件名
      const subject = `SI製品 物件手配品の見積依頼　${subjectValue}`;
      
      // メール本文
      const body = `嶋澤さん\n\n`
                 + `お疲れ様です、木伏です。\n`
                 + `下記についてよろしくお願いします。\n\n`
                 + `件名：\n`
                 + `${subjectValue}\n\n`
                 + `コメント内容：\n`
                 + `${content}`;

      // メール送信
      MailApp.sendEmail(TARGET_EMAIL, subject, body);
      console.log(`メール送信完了: 行=${row}, 件名=${subjectValue}`);
    }
  });

  // チェック時間を更新
  props.setProperty('LAST_CHECK_TIME', now.toISOString());
}
