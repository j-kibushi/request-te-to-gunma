/**
 * 見積回答があった日時を記入する。
*/
function onEditday(e) {
  // 編集されたシートと範囲を取得
  var sheet = e.source.getActiveSheet();
  var editedRange = e.range;
  
  // 編集されたシートが「フォームの回答 1」であるか確認
  if (sheet.getName() == '依頼一覧') {
    // 編集された列がf列（列番号6）であるか確認
    if (editedRange.getColumn() == 6) {
      // 編集された行番号を取得
      var row = editedRange.getRow();
      
      // E列に値があるか確認
      var value = editedRange.getValue();
      if (value != "") {
        // H列（列番号8）に現在の日付と時刻を書き込む
        sheet.getRange(row, 8).setValue(new Date());
      }
    }
  }
}