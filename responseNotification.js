/**
 * 回答が記入されたことを通知する。
*/

function onEditnew(e) {
  var sheet = e.source.getActiveSheet();
  var editedRange = e.range;

  // 編集されたシートが「依頼一覧」であるか確認
  if (sheet.getName() == '依頼一覧') {
    // 編集された列がF列（列番号6）であるか確認
    if (editedRange.getColumn() == 6) {
      // 編集された行番号を取得
      var row = editedRange.getRow();
      
      // E列に値があるか確認
      var value = editedRange.getValue();
      if (value != "") {
        // G列に日付と時刻を書き込む
        sheet.getRange(row, 7).setValue(new Date());

        // メール送信の準備
        var recipient = "j_kibushi@carecom.co.jp";
        var subject = "「依頼一覧」シートに新しい回答が入力されました";
        var body = "「依頼一覧」シートのE列に新しい回答が入力されました。\n\n"
                 + "編集されたセル: " + editedRange.getA1Notation() + "\n"
                 + "入力された値: " + value + "\n"
                 + "スプレッドシートのURL: " + e.source.getUrl();

        // メールを送信
        MailApp.sendEmail(recipient, subject, body);
      }
    }
  }
}