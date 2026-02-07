/**
 * 群馬工場への見積依頼メールを送るためのスクリプトです。
 * Googleフォームの回答が送信されたときに自動実行され、内容に応じたメールを送信します。
 */
function sendMailOnFormSubmit(e) {
  // メールの宛先
  const RECIPIENT = "m_shimazawa@carecom.co.jp,h_arai@carecom.co.jp";

  // フォームの質問項目名（スプレッドシートのヘッダー名）
  // ログで確認した正しいキーに修正済みの状態です。
  const COL_SUBJECT = "件名";
  const COL_NOTES_LINK = "見積依頼WF　  NotesURLリンク"; 
  const COL_DETAILS = "見積内容";
  const COL_PROCESS = "工程";

  // フォームの回答データを質問項目名をキーにして取得します。
  const values = e.namedValues;
  const subjectFromForm = values[COL_SUBJECT] ? values[COL_SUBJECT][0] : "";
  const notesLink = values[COL_NOTES_LINK] ? values[COL_NOTES_LINK][0] : "";
  const estimateDetails = values[COL_DETAILS] ? values[COL_DETAILS][0] : "";
  const process = values[COL_PROCESS] ? values[COL_PROCESS][0] : "";

  let mailSubject = "";
  let mailBody = "";

  // 「工程」(F列) の値に応じて、メールの件名と本文を作成します。
  switch (process) {
    case "見積":
      mailSubject = `群馬へのSIハード見積依頼　 ${subjectFromForm}`;
      mailBody = `To　嶋澤さん、新井さん

お疲れ様です、基盤展開TE　木伏です。
下記手配品の見積書取得をお願いいたします。

${estimateDetails}

＜見積依頼WFへのリンク＞
${notesLink}


【見積書を取得後】
下記URLの「見積書URL」列へ見積書PDFのURL、「連絡事項」列へメモを入力してください。
https://docs.google.com/spreadsheets/d/1B7WqcJitpAv34qip8t6sOMBPK0sdvtHdBs7z54TBjxk/edit?usp=sharing
`;
      break;

    case "受注":
      mailSubject = `【受注】　 ${subjectFromForm}`;
      mailBody = `TO　嶋澤さん、新井さん

お疲れ様です、木伏です。
見積書の取得をお願いします。受注案件ですので早めの回答希望です。

${estimateDetails}


物件名：${subjectFromForm}
見積依頼WF：
${notesLink}

【見積書を取得後】
下記URLの「見積書URL」列へ見積書PDFのURL、「連絡事項」列へメモを入力してください。
https://docs.google.com/spreadsheets/d/1B7WqcJitpAv34qip8t6sOMBPK0sdvtHdBs7z54TBjxk/edit?usp=sharing
`;
      break;

    default:
      console.log(`工程が「${process}」のため、メールは送信されませんでした。`);
      return;
  }

  // メールを送信します。
  try {
    MailApp.sendEmail({
      to: RECIPIENT,
      subject: mailSubject,
      body: mailBody
    });
    console.log(`メールが正常に送信されました。宛先: ${RECIPIENT}, 件名: ${mailSubject}`);
  } catch (error) {
    console.error(`メールの送信に失敗しました。エラー: ${error.toString()}`);
  }
}
