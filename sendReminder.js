/**
 * 見積書の催促、進捗確認のメールを送る。
*/

function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('リマインド')
      .addItem('進捗確認メールを送信', 'sendReminderEmail')
      .addToUi();
}

function sendReminderEmail() {
  // メッセージボックスを表示してユーザーに行選択を促す
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '進捗確認メールを送信',
    '進捗確認を行いたい依頼の行を選択してください。',
    ui.ButtonSet.OK_CANCEL
  );

  // ユーザーが「OK」ボタンを押さなかった場合、処理を中断
  if (response !== ui.Button.OK) {
    return;
  }
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  if (sheet.getName() !== '依頼一覧') {
    Browser.msgBox('この機能は「依頼一覧」シートでのみ利用できます。');
    return;
  }
  
  const activeRange = sheet.getActiveRange();
  const startRow = activeRange.getRow();
  
  if (startRow <= 1) {
    Browser.msgBox('ヘッダー行は選択できません。');
    return;
  }
  
  const rowData = sheet.getRange(startRow, 1, 1, 4).getValues()[0];
  const subjectText = rowData[1];
  const valueB = rowData[2];
  const valueC = rowData[3];
  
  // E列（5列目）から担当営業名を取得
  const salesRepName = sheet.getRange(startRow, 5).getValue();
  
  // メールアドレスを設定
  const toRecipients = "m_shimazawa@carecom.co.jp,h_arai@carecom.co.jp";
  let ccRecipients = "j_kibushi@carecom.co.jp,t_oonishi@carecom.co.jp"; 

  // 担当営業（G列）の情報を取得してCCに追加
  if (salesRepName) {
    try {
      const salesMasterSsId = "104S53Ag64cg-dfcgfC3V_dO8LlP9auFA_XRBAaE-GwY";
      const salesMasterSheet = SpreadsheetApp.openById(salesMasterSsId).getSheetByName("セールスマスタ");
      const data = salesMasterSheet.getDataRange().getValues();
      
      // セールスマスタから氏名(A列)が一致する行のメールアドレス(B列)を検索
      const targetRow = data.find(r => r[0] === salesRepName);
      if (targetRow && targetRow[1]) {
        ccRecipients += "," + targetRow[1];
      }
    } catch (err) {
      console.warn(`担当営業のメールアドレス取得に失敗しました。エラー: ${err.toString()}`);
    }
  }
  
  const subject = "【確認】見積書取得依頼の進捗状況について " + subjectText;
  
  const body = "先日依頼いたしました下記物件について、進捗状況の確認がありました。\n\n"
             + "すみませんがご確認を頂き、急ぎ回答がいただけるようフォローアップよろしくお願いします。\n\n"
             + "--------------------------------------------------\n"
             + "件名: " + subjectText + "\n"
             + "見積依頼WF URL: " + valueB + "\n"
             + "依頼内容: "+ "\n" + valueC + "\n"
             + "--------------------------------------------------\n\n"
             + "SI製品　物件手配品の見積依頼（基盤展開TE→群馬）: " + SpreadsheetApp.getActiveSpreadsheet().getUrl();
             
  // メールを送信
  MailApp.sendEmail({
    to: toRecipients,
    subject: subject,
    body: body,
    cc: ccRecipients
  });

  // メール送信日時をI列に記録
  const now = new Date();
  sheet.getRange(startRow, 10).setValue(now);
  
  Browser.msgBox('進捗確認のメールが送信されました。');
}