/**
 * @OnlyCurrentDoc
 */

//スプレッドシートの読み込み
var sheet       = SpreadsheetApp.getActiveSheet();
  
//スプレッドシートの項目の変数化-----------------------------------------------------------------------------
var date        = sheet.getRange(2, 1, sheet.getLastRow(), sheet.getLastColumn());  //"証明書期限"
var url         = sheet.getRange(2, 2, sheet.getLastRow(), sheet.getLastColumn());  //"コモンネーム"
var key_length  = sheet.getRange(2, 3, sheet.getLastRow(), sheet.getLastColumn());  //"鍵長"
var poodle      = sheet.getRange(2, 4, sheet.getLastRow(), sheet.getLastColumn());  //"脆弱性"
var terminal    = sheet.getRange(2, 5, sheet.getLastRow(), sheet.getLastColumn());  //"対象端末"
var middle      = sheet.getRange(2, 6, sheet.getLastRow(), sheet.getLastColumn());  //"中間証明書情報"
var period      = sheet.getRange(2, 7, sheet.getLastRow(), sheet.getLastColumn());  //"期間"
var quantum     = sheet.getRange(2, 8, sheet.getLastRow(), sheet.getLastColumn());  //"数量"
var company     = sheet.getRange(2, 9, sheet.getLastRow(), sheet.getLastColumn());  //"発行会社"
var last        = sheet.getRange(2, 10, sheet.getLastRow(), sheet.getLastColumn()); //"前回作業者"
var worker      = sheet.getRange(2, 11, sheet.getLastRow(), sheet.getLastColumn()); //"作業者"
var steps       = sheet.getRange(2, 12, sheet.getLastRow(), sheet.getLastColumn()); //"発行手続者"
var base        = sheet.getRange(2, 13, sheet.getLastRow(), sheet.getLastColumn()); //"運用基盤"
var web_site    = sheet.getRange(2, 14, sheet.getLastRow(), sheet.getLastColumn()); //"用途・サイト名"
var unit        = sheet.getRange(2, 15, sheet.getLastRow(), sheet.getLastColumn()); //"対象ユニット"
var person      = sheet.getRange(2, 16, sheet.getLastRow(), sheet.getLastColumn()); //"向き合い担当者"
var person_mail = sheet.getRange(2, 17, sheet.getLastRow(), sheet.getLastColumn()); //"向き合い担当者アドレス"
var corporation = sheet.getRange(2, 18, sheet.getLastRow(), sheet.getLastColumn()); //"申請法人"
var post        = sheet.getRange(2, 19, sheet.getLastRow(), sheet.getLastColumn()); //"申請部署"
var etc         = sheet.getRange(2, 20, sheet.getLastRow(), sheet.getLastColumn()); //"備考"
var send_list   = sheet.getRange(2, 21, sheet.getLastRow() - 1, 1);                 //"第一次メール送信済みフラグ"
var hojin       = sheet.getRange(2, 22, sheet.getLastRow(), sheet.getLastColumn()); //"申請法人情報"
var sekininsya  = sheet.getRange(2, 23, sheet.getLastRow(), sheet.getLastColumn()); //"申請責任者情報"
var csr         = sheet.getRange(2, 24, sheet.getLastRow(), sheet.getLastColumn()); //"CSR情報"
var second_list = sheet.getRange(2, 25, sheet.getLastRow() - 1, 1);                 //"第二次メール送信済みフラグ"
var file        = sheet.getRange(2, 26, sheet.getLastRow(), sheet.getLastColumn()); //"見積書"
//-------------------------------------------------------------------------------------------------------------

//配列格納-------------------------------------------
var date_box        = date.getValues();
var url_box         = url.getValues();
var key_length_box  = key_length.getValues();
var poodle_box      = poodle.getValues();
var terminal_box    = terminal.getValues();
var middle_box      = middle.getValues();
var period_box      = period.getValues();
var quantum_box     = quantum.getValues();
var company_box     = company.getValues();
var last_box        = last.getValues();
var worker_box      = worker.getValues();
var steps_box       = steps.getValues();
var base_box        = base.getValues();
var web_site_box    = web_site.getValues();
var unit_box        = unit.getValues();
var person_box      = person.getValues();
var person_mail_box = person_mail.getValues();
var corporation_box = corporation.getValues();
var post_box        = post.getValues();
var etc_box         = etc.getValues();
var send_list_box   = send_list.getValues();
var hojin_box       = hojin.getValues();
var sekininsya_box  = sekininsya.getValues();
var csr_box         = csr.getValues();
var second_list_box = second_list.getValues();
var file_box        = file.getValues();
//----------------------------------------------------
  

//第一次メール送信
function sendMail(){
  //入力されているデータ(行)分ループ
  for(var i = 0; i < date_box.length - 1; ++i){
    //送信先メールアドレスの設定
    var address     = "TO宛のメールアドレス";
    var cc_address  = "CC宛のメールアドレス";
    //担当者のメールアドレスの追加
    address += "," + person_mail_box[i][0];
    Logger.log(address);
    //スプレッドシート内"証明書期限"を参照
    var getDate = date_box[i][0];
    Logger.log(getDate);
    //現在の時刻を取得
    var nowDate = new Date();
    Logger.log(nowDate);
    //証明書期限と現在の時刻の差が30日前か比較
    var margin　= (getDate.getTime()-(30*24*60*60*1000) < nowDate.getTime());
    //証明書期限が30日前の場合
    if(margin == true){
      //自動メール送信有無が空の場合
      if(send_list_box[i][0] == ""){
        //メール送信 & メール本文内容
        MailApp.sendEmail(  
          address,
          "【SSL証明書更新期限通知】" + url_box[i][0],
          "ご担当者様" + String.fromCharCode(10) + 
          String.fromCharCode(10) + 
          "SSL証明書の期限が1ヶ月前に迫っています。" + String.fromCharCode(10) + 
          "詳細のご連絡は後ほど致しますので更新是非の確認を宜しくお願い致します。" + 
          String.fromCharCode(10) +
          String.fromCharCode(10) +
          "---SSL証明書情報---" + String.fromCharCode(10) + 
          "コモンネーム" + "：" + url_box[i][0] + String.fromCharCode(10) + 
          "証明書期限" + "：" + date_box[i][0] + String.fromCharCode(10) + 
          "鍵長" + "：" + key_length_box[i][0] + String.fromCharCode(10) + 
          "脆弱性" + "：" + poodle_box[i][0] + String.fromCharCode(10) + 
          "対象端末" + "：" + terminal_box[i][0] + String.fromCharCode(10) + 
          "中間証明書情報" + "：" + middle_box[i][0] + String.fromCharCode(10) + 
          "期間" + "：" + period_box[i][0] + String.fromCharCode(10) + 
          "数量" + "：" + quantum_box[i][0] + String.fromCharCode(10) + 
          "発行会社" + "：" + company_box[i][0] + String.fromCharCode(10) + 
          "前回作業者" + "：" + last_box[i][0] + String.fromCharCode(10) + 
          "作業者" + "：" + worker_box[i][0] + String.fromCharCode(10) + 
          "発行手続者" + "：" + steps_box[i][0] + String.fromCharCode(10) + 
          "運用基盤" + "：" + base_box[i][0] + String.fromCharCode(10) + 
          "用途・サイト名" + "：" + web_site_box[i][0] + String.fromCharCode(10) + 
          "対象ユニット" + "：" + unit_box[i][0] + String.fromCharCode(10) + 
          "向き合い担当者" + "：" + person_box[i][0] + String.fromCharCode(10) + 
          "申請法人" + "：" + corporation_box[i][0] + String.fromCharCode(10) + 
          "申請部署" + "：" + post_box[i][0] + String.fromCharCode(10) +
          String.fromCharCode(10) + 
          String.fromCharCode(10) +
          "※このメールは自動送信しています。不備が御座いましたら担当者にご連絡頂けますと幸いです。", {cc:cc_address}) 
        var send_list_flag = sheet.getRange(i+2, 21);
        send_list_flag.setValue("◯");
        //送信先メールアドレスの初期化
        delete address;
        var send_url  = url_box[i][0];
        var send_date = date_box[i][0];
        sendHipchat(send_url, send_date);
        Logger.log("第一次メール送信しました")
      }else if(send_list_box[i][0] == "◯" && second_list_box[i][0] == ""){
        // 添付ファイル用の配列を作成
        var attachmentFiles = new Array();
        
        // 添付ファイルを取得
        var attachment_Id = file_box[i][0];
        var attachment;
        var attachment_URL = "";
        
        if (attachment_Id != "") {
          // Google Driveから添付ファイルのデータを取得
          var attachment = DriveApp.getFileById(attachment_Id).getBlob();
          if (attachment != null) {
            // Gmail添付用のデータを作成（ファイル名、mimeタイプ、バイト配列を指定）
            attachmentFiles.push({fileName:attachment.getName(), mimeType: attachment.getContentType(), content:attachment.getBytes()});
            attachment_URL = DriveApp.getFileById(attachment_Id).getUrl();
          }
        
          MailApp.sendEmail(  
          address,
          "【SSL証明書更新通知】" + url_box[i][0],
          "ご担当者様" + String.fromCharCode(10) + 
          String.fromCharCode(10) + 
          "下記SSL証明書の有効期限が迫っております。" + String.fromCharCode(10) + 
          String.fromCharCode(10) + 
          "---SSL証明書情報---" + String.fromCharCode(10) + 
          "用途・サイト名" + "：" + web_site_box[i][0] + String.fromCharCode(10) + 
          "コモンネーム" + "：" + url_box[i][0] + String.fromCharCode(10) + 
          "数量" + "：" + quantum_box[i][0] + String.fromCharCode(10) + 
          "鍵長" + "：" + key_length_box[i][0] + String.fromCharCode(10) +
          "発行会社" + "：" + company_box[i][0] + String.fromCharCode(10) +
          "証明書期限" + "：" + date_box[i][0] + String.fromCharCode(10) +
          String.fromCharCode(10) +
          String.fromCharCode(10) +                                   
          "つきましては更新の有無をご確認頂き、更新する場合は下記の申請情報に変更がないかご確認下さい。" + 
          String.fromCharCode(10) +
          String.fromCharCode(10) +
          "申請法人情報" + "----------" + String.fromCharCode(10) +
          hojin_box[i][0] + String.fromCharCode(10) + 
          String.fromCharCode(10) + 
          "申請責任者情報" + "----------" + String.fromCharCode(10) +
          sekininsya_box[i][0] + String.fromCharCode(10) + 
          String.fromCharCode(10) +
          "CSR情報" + "----------" + String.fromCharCode(10) +
          csr_box[i][0] + String.fromCharCode(10) +
          String.fromCharCode(10) +
          String.fromCharCode(10) +
          "お見積りを添付いたしますのでご確認を宜しくお願い致します。"  + String.fromCharCode(10) +
          String.fromCharCode(10) +
          String.fromCharCode(10) +
          "証明書の形式がSHA-1からSHA-2に切り替わる方針になっております。一部古いガラケー端末では表示できなくなります。ご確認ください。" + String.fromCharCode(10) +
          String.fromCharCode(10) +
          String.fromCharCode(10) +
          "※このメールは自動送信しています。不備が御座いましたら担当者にご連絡頂けますと幸いです。", {cc:cc_address, attachments:attachmentFiles})
          var second_list_flag = sheet.getRange(i+2, 25);
          second_list_flag.setValue("◯");
          delete address;
          Logger.log("第一次メール送信済み")
          Logger.log("第二次メール送信しました")
        }else{
          MailApp.sendEmail(  
          address,
          "【SSL証明書更新通知】" + url_box[i][0],
          "ご担当者様" + String.fromCharCode(10) + 
          String.fromCharCode(10) + 
          "下記SSL証明書の有効期限が迫っております。" + String.fromCharCode(10) + 
          String.fromCharCode(10) + 
          "---SSL証明書情報---" + String.fromCharCode(10) + 
          "用途・サイト名" + "：" + web_site_box[i][0] + String.fromCharCode(10) + 
          "コモンネーム" + "：" + url_box[i][0] + String.fromCharCode(10) + 
          "数量" + "：" + quantum_box[i][0] + String.fromCharCode(10) + 
          "鍵長" + "：" + key_length_box[i][0] + String.fromCharCode(10) +
          "発行会社" + "：" + company_box[i][0] + String.fromCharCode(10) +
          "証明書期限" + "：" + date_box[i][0] + String.fromCharCode(10) +
          String.fromCharCode(10) +
          String.fromCharCode(10) +                                   
          "つきましては更新の有無をご確認頂き、更新する場合は下記の申請情報に変更がないかご確認下さい。" + 
          String.fromCharCode(10) +
          String.fromCharCode(10) +
          "申請法人情報" + "----------" + String.fromCharCode(10) +
          hojin_box[i][0] + String.fromCharCode(10) + 
          String.fromCharCode(10) + 
          "申請責任者情報" + "----------" + String.fromCharCode(10) +
          sekininsya_box[i][0] + String.fromCharCode(10) + 
          String.fromCharCode(10) +
          "CSR情報" + "----------" + String.fromCharCode(10) +
          csr_box[i][0] + String.fromCharCode(10) +
          String.fromCharCode(10) +
          String.fromCharCode(10) +
          "証明書の形式がSHA-1からSHA-2に切り替わる方針になっております。一部古いガラケー端末では表示できなくなります。ご確認ください。" + String.fromCharCode(10) +
          String.fromCharCode(10) +
          String.fromCharCode(10) +
          "※このメールは自動送信しています。不備が御座いましたら担当者にご連絡頂けますと幸いです。", {cc:cc_address})
          var second_list_flag = sheet.getRange(i+2, 25);
          second_list_flag.setValue("◯");
          delete address;
          Logger.log("第一次メール送信済み")
          Logger.log("第二次メール送信しました(見積もり書なし)")
        }
      }else if(send_list_box[i][0] == "◯" && second_list_box[i][0] == "◯"){  
        Logger.log("第二次メール送信済み")
      }else{
        Looger.log("error発生。スプレッドシートに不備はありませんか？")
      }
    }else{
      Logger.log("まだ期限が迫ってきていません")
    }
  }
}

function sendHipchat(send_url, send_date){
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var url = doc.getUrl();

  //var room_id = "HipchatテストルームID" // APIテスト
  var room_id = "HipchatルームID" // 本番用
  var message = "以下のSSL証明書の更新期限が30日前です。";
  var hipchat =
      {
        "auth_token" : "任意のauth_token",
        "room_id" : room_id,
        "from": "Hipchat上の表示名",
        "color": '通知時の色',
        "format" : "フォーマット",
        "message_format" : "メッセージフォーマット",
        "message" : message + String.fromCharCode(10) + send_url + String.fromCharCode(10) + send_date
      };

  var options =
      {
        "method" : "post",
        "payload" : hipchat
      };
  UrlFetchApp.fetch('https://api.hipchat.com/v1/rooms/message', options);
}