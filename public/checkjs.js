//小寫轉大寫
function toUpperCase(cthis) {
    cthis.value = cthis.value.toUpperCase();
}

//大寫轉小寫
function toLowerCase(cthis) {
    cthis.value = cthis.value.toLowerCase();
}

//檢查空白值
function IsEmpty(cthis) {
    if (cthis.value == '') {
        alert('欄位值不能為空白');
    }
}

//檢查YYYY/MM/DD日期格式
function CheckDate(cthis) {
    if (cthis.value == '') { return true; }
    var re = new RegExp("^([0-9]{4})[./]{1}([0-9]{1,2})[./]{1}([0-9]{1,2})$");
    var strDataValue
    var str = cthis.value;
    var infoValidation = true;
    
    if ((strDataValue = re.exec(str)) != null){
        var i;
        i = parseFloat(strDataValue[1]);
        if (i <= 0 || i > 9999){ // 年
            infoValidation = false;
        }
        i = parseFloat(strDataValue[2]);
        if (i <= 0 || i > 12){ // 月
            infoValidation = false;
        }
        i = parseFloat(strDataValue[3]);
        if (i <= 0 || i > 31){ // 日
            infoValidation = false;
        }
    }else{
        infoValidation = false;
    }

    if (!infoValidation){
        alert('請輸入 YYYY/MM/DD 日期格式');
    }
    return infoValidation;
}

//檢查資料是否為數字型態內容
function IsNumeric(cthis) {
    var strValidChars = "0123456789.-";
    var strChar;
    var blnResult = true;
    var strString = cthis.value;
    if (strString.length == 0) return false;
    for (i = 0; i < strString.length && blnResult == true; i++) {
        strChar = strString.charAt(i);
        if (strValidChars.indexOf(strChar) == -1) {
            blnResult = false;
        }
    }
    if (blnResult == false) {
        alert('欄位內容限定輸入數字格式');
    }
    return blnResult;
}

//限制輸入資料內容
function SetTextMasK(cthis,climitstr) {
    var strValidChars = climitstr;
    var strChar;
    var blnResult = true;
    var strString = cthis.value;
    if (strString.length == 0) return false;
    for (i = 0; i < strString.length && blnResult == true; i++) {
        strChar = strString.charAt(i);
        if (strValidChars.indexOf(strChar) == -1) {
            blnResult = false;
        }
    }
    if (blnResult == false) {
        alert('欄位輸入格式不正確');
    }
    return blnResult;
}

//取得2日期之間的差值，僅限於"日"單位
function DateDiff(asStartDate, asEndDate) {
    var miStart = Date.parse(asStartDate.replace(/\-/g, '/'));
    var miEnd = Date.parse(asEndDate.replace(/\-/g, '/'));
    return (miEnd - miStart) / (1000 * 24 * 3600);
}
