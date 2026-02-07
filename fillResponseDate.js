/**
 * 見積回答があった日時を記入する
*/
function onEditday(e) {
  // 編集されたシートと範囲を取得
  var sheet = e.source.getActiveSheet();
  var editedRange = e.range;
  
  // 編集されたシートが「依頼一覧」であるか確認
  if (sheet.getName() == '依頼一覧') {
    // 編集された列がH列（列番号8）であるか確認
    if (editedRange.getColumn() == 8) {
      // 編集された行番号を取得
      var row = editedRange.getRow();
      
      // H列に値があるか確認
      var value = editedRange.getValue();
      if (value != "") {
        // J列（列番号10）に現在の日付と時刻を書き込む
        sheet.getRange(row, 10).setValue(new Date());
      }
    }
  }
}