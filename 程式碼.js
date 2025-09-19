function myFunction() {
  
}

/**
 * 當 Google 表單提交時觸發此函式。
 * 它會向提交表單的人發送一封確認電子郵件。
 *
 * @param {Object} e 來自表單提交觸發器的事件物件。
 */
function onFormSubmit(e) {
  // 從事件物件中獲取表單回應物件。
  const formResponse = e.response;

  // 從表單回應中獲取所有項目回應。
  const itemResponses = formResponse.getItemResponses();

  // 創建一個空物件來存儲我們的表單數據。
  const formData = {};

  // 遍歷每個項目回應以獲取問題標題和答案。
  for (let i = 0; i < itemResponses.length; i++) {
    const itemResponse = itemResponses[i];
    const question = itemResponse.getItem().getTitle();
    const answer = itemResponse.getResponse();
    formData[question] = answer;
  }

  // 假設您的表單中有一個名為「電子郵件」的欄位。
  // 請務必將 '電子郵件' 替換為您表單中收集電子郵件地址的實際問題標題。
  const email = formData['電子郵件'];

  // 如果找不到電子郵件地址，則記錄錯誤並且不執行任何操作。
  if (!email) {
    Logger.log("找不到電子郵件地址。無法發送電子郵件。");
    return;
  }

  // 定義電子郵件的主旨和內文。
  const subject = "感謝您的提交！";
  const message = "您好，\n\n感謝您提交表單。我們已收到您的回應。\n\n祝好，\n團隊";

  // 發送電子郵件。
  MailApp.sendEmail(email, subject, message);

  Logger.log("已將確認電子郵件發送至：" + email);
}