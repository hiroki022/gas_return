function doGet(e){
  var htmloutput = HtmlService.createTemplateFromFile("comp").evaluate();
  return htmloutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e){
  var number=e.parameter.number;
  comp(number);
  var htmloutput = HtmlService.createTemplateFromFile("完了").evaluate();
  return htmloutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function comp(number) {
  //有効なGooglesプレッドシートを開く
  var ss = SpreadsheetApp.openById('1gMn6N77DZW4j9h7eRhI7jr3pIK4EzG0IW_mGlf3YzOQ');
  var sheet = ss.getSheetByName('フォーム')
  var cam_list = ss.getSheetByName('カメラの個数');
  
  //新規予約された行番号を取得
  var num_row = sheet.getLastRow();
  //「カメラの個数」シートの一番下の行番号を取得
  var num_row2 = cam_list.getLastRow();
  //学籍番号を配列に格納
  var list_number = sheet.getRange(2,10,num_row,1).getValues();
  
  //test用
  //number = "s1701078";
  Logger.log(number);
  Logger.log(list_number[0]);
  
  var cals = CalendarApp.getCalendarById("ug50e7guqcnmpog1eest35dmk0@group.calendar.google.com");
  
  var thing = "返却が完了しました";
  var name;
  var mail;
  
  var rental
  var j = 2;
  while(j <= num_row)
  {
    rental = sheet.getRange(j, 9).getValue();
    if(rental == "未返却")
    {
      if(list_number[j-2] == number)
      {
        sheet.getRange(j,9).setValue("返却済み");
        name = sheet.getRange(j, 3).getValue();
        mail = sheet.getRange(j,2).getValue(); //部員への完了を知らせるメール
        MailApp.sendEmail(mail,name + "さんのカメラ返却完了のお知らせ",thing); //メールを送信
        name = sheet.getRange(j, 3).getValue();
        //mail = cam_list.getRange(2,7).getValue(); //部長への完了を知らせるメール
        //MailApp.sendEmail(mail,name + "さんのカメラ返却完了のお知らせ",thing); //メールを送信
        
        //借りたいカメラの名前　F列
        var rental_camera = sheet.getRange(num_row,6).getValue();
        
        //今あるカメラのリスト
        var list = cam_list.getRange(1,1,num_row2,1).getValues();
        
        //その人が借りたいカメラ
        //リストの中に借りたいカメラはどの行か調べる
        var i = 0;
        
        while( i < rental_camera + 1 )
        {
          if(list[i-1] == rental_camera)
          {
            Logger.log('一致しました');
            break;
          }
          else
          {
            i = i +1;
          }
        }
        
        //カレンダーから予定を削除
        var nname = sheet.getRange(j,3).getValue();
        var camera = cam_list.getRange(i,2).getDisplayValue();
        var thing = nname+"様　"+camera+"　のご予約"
        var ndates = sheet.getRange(j,7).getValue();
        var ndatee = sheet.getRange(j,8).getValue();
        var events = cals.getEvents(ndates,ndatee);
        if(events[0]　== null) {
          break;
        }
        else{
          events[0].deleteEvent();
        }
      }
    }
    j = j + 1;
  }
  //"貸し出し状況"に-1をする
  var rent = cam_list.getRange(i,4).getValue();
  cam_list.getRange(i,4).setValue(rent - 1);
  if(rent <= 0)
  {
  cam_list.getRange(i,4).setValue(0);
  }
  
}