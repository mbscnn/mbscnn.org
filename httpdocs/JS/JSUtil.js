﻿// JavaScript Document
// By Titan
//************************************************************
// 目的：去除前後空白
// 傳回：
//************************************************************
String.prototype.trim = function trim() {
	return this.replace(/^\s+|\s+$/g, "");
}

//************************************************************
// 目的：檢查是否為正確的日期（民國年）
// 傳回：true (valid), false (invalid)
//************************************************************
function isValidTWDate(str){
	str=str.trim().replaceStr(str,"/","",2);//來源字串:String, 原始字串:String, 目標字串:String[, 取代數目:Number
	str=str.trim().replaceStr(str,".","",2);
	//alert(str);
	var len = str.length;
	if (len == 6) {
		return isValidCDate( "0" + str.substr(0,2),str.substr(2,2),str.substr(4,2));
	} else if (len == 7) {
		return isValidCDate(str.substr(0,3),str.substr(3,2),str.substr(5,2));
	} 	
	return false;
}

function isValidTWDate1(str){
	var len = str.value.length;
	if (len == 0)
	{
	return true;
	}
	
	if (len == 6) {
		if (isValidCDate( "0" + str.value.substr(0,2),str.value.substr(2,2),str.value.substr(4,2)) == true)
		{
		return true;
		}
	} else if (len == 7) {
		if (isValidCDate(str.value.substr(0,3),str.value.substr(3,2),str.value.substr(5,2)) == true)
		{
		return true;
		}
	} 	
	alert("日期格式不正確！");
	window.event.srcElement.select();
	window.event.srcElement.focus();
}

function isValidCDate(strYear, strMonth, strDay){
		 
	var intYear = parseInt(strYear, 10) + 1911;
	strYear = "" + intYear;
	return isValidDate(strYear, strMonth, strDay);
}

//************************************************************
// 目的：檢查是否為正確的日期（西元年）
// 傳回：true (valid), false (invalid)
//************************************************************
function isValidDate(strYear, strMonth, strDay){
 	// empty string
	if (isEmptyStr(strYear) || isEmptyStr(strMonth) || isEmptyStr(strDay)){
		return false;
	}

	// numeric
	if (isNaN(strYear) || isNaN(strMonth) || isNaN(strDay)){
		return false;
	}
	var intYear = parseInt(strYear, 10);
	var intMonth = parseInt(strMonth, 10);
	var intDay = parseInt(strDay, 10);

	if (intYear < 1900 || intYear > 2200){
		return false;
	}

	if (!intRange(intMonth, 1, 12)) {
		return false;
	}

	var aryMonthDay = new Array(12);

	aryMonthDay[0] = 31; aryMonthDay[1] =  (isLeapYear(strYear) ? 29 : 28); aryMonthDay[2] = 31; aryMonthDay[3] = 30;
	aryMonthDay[4] = 31; aryMonthDay[5] = 30; aryMonthDay[6] = 31; aryMonthDay[7] = 31;
	aryMonthDay[8] = 30; aryMonthDay[9] = 31; aryMonthDay[10] = 30; aryMonthDay[11] = 31;

	if (!intRange(intDay, 1, aryMonthDay[intMonth-1])) {
		return false;
	}

	return true;
}

//************************************************************
// 目的：檢查是否為潤年
// 傳回：true (潤年), false (非潤年)
//************************************************************
function isLeapYear(strYear){

	var blnLeapYear;
	if ((strYear % 400)  == 0 ) {
		blnLeapYear = true;
	}else if((strYear % 4) == 0) {
		if((strYear % 100) == 0) {
			blnLeapYear = false
		}
		else {
			blnLeapYear = true;
		}
	}else {
		blnLeapYear = false;
	}
	return blnLeapYear;
}

//************************************************************
// 目的：檢查是否有效範圍的整數
// 傳回：true (in range), false (out of range)
//************************************************************
function intRange(strInput, intLow, intUpper){

	var intNum = parseInt(stripCommas(strInput), 10);

	if ((intNum < intLow) || (intNum > intUpper)) {
		return false;
	}else {
		return true;
	}
}

//************************************************************
// 目的：檢查是否有效範圍的浮點數，有效範圍請輸入浮點數
// floatRange(88.1,0.00,100.01)
// 傳回：true (in range), false (out of range)
//************************************************************
function floatRange(strInput, floatLow, floatUpper){

	var floatNum = parseFloat(stripCommas(strInput));

	if ((floatNum < floatLow) || (floatNum > floatUpper)){
		return false;
	}else {
	    return true;
    }
}

//************************************************************
// 目的：將","取代為空字串  1,234 --> 1234
// 傳回：
//************************************************************
function stripCommas(numString){
	var re = /,/g;
	if (typeof numString == "number"){
		return numString;
	}
	return numString.replace(re,"");
}

//************************************************************
// 目的：將"%"取代為空字串  24% --> 24
// 傳回：
//************************************************************
function stripPercent(numString){		
	var re = /%/g;	
	if (typeof numString == "number"){
		return numString;
	}
	return numString.replace(re,"");
}

//************************************************************
// 目的：檢查是否為數字
// 傳回：true (valid), false (invalid)
//************************************************************
function isNumber(strNum){

	if (isEmptyStr(strNum)){
		return false;
	}

	var num = new Number(stripCommas(strNum));
	if (isNaN(num)){
		return false;
	}
	return true;
}

//************************************************************
// 目的：檢查字串是否為空
// 傳回：
//************************************************************
function isEmptyStr(str){

	var len = str.length;
	var i = 0;
	var isEmpty = true;
	for (i=0; (i < len) && (isEmpty); i++){
		isEmpty = (str.charAt(i) == " ");
	}
	return isEmpty;
}

//************************************************************
// 目的：檢查是否為數字,如果不是將textbox內容清空
//		intLow-->下限,intUpper-->上限
// 傳回：
//************************************************************
function checkInt(obj, intLow, intUpper){

	var sValue=obj.value;
	if (sValue=="")
	    return true;
 	if(!isNumber(sValue) || !intRange(sValue,intLow,intUpper)){
 		obj.value="";		
		return false;
	}else{
		this.document.all(obj.id).value=obj.value;
		return true;
	}
}

//************************************************************
// 目的：檢查是否為數字,如果不是將textbox內容清空
//		floatLow-->下限,floatUpper-->上限
// 傳回：
//************************************************************
function checkfloat(obj, floatLow, floatUpper){
	 
	var sValue=obj.value;

	if(!isNumber(sValue) || !floatRange(sValue,floatLow,floatUpper)){
		obj.value="";
	}else{
		this.document.all(obj.id).value=obj.value;
	}
}

//************************************************************
// 目的：判斷下拉選單是否有被選取，如果沒有選取，出現提示訊息
// 傳回：Boolean
//************************************************************
function isChosen(selectObjsct,sMessage) {
	if (selectObjsct !=null){
		if (selectObjsct.selectedIndex == 0) {
			if(sMessage != null && sMessage.length > 0){
				alert(sMessage);
			}
			return false;
		} else {
			return true;
		}
	}
	return true;
}


//************************************************************
// 目的：判斷Radio But是否有被選取，如果沒有選取，出現提示訊息
// 傳回：Boolean
//************************************************************
function isValidRadio(radioObject,sMessage) {
	var valid = false;

	if (radioObject !=null){
		var element=window.document.getElementsByName(radioObject.name);

		if (element.length==1){
			if (element[0].checked){
				return true;
			}
		}else{
			for (var i = 0; i < radioObject.length; i++) {
				if (radioObject[i].checked){
					return true;
				}
			}
		}

		if(sMessage != null && sMessage.length > 0){
			alert(sMessage);
			return false;
		}
	}
	return true;
}

//************************************************************
// 目的：判斷TexBoxt是否有為空，如果為空，出現提示訊息
// 傳回：Boolean
//************************************************************
function isNotEmpty(elem,sMessage){
	if (elem !=null) {
		var str = elem.value;
		
		if(str == null || str.length == 0) {
			if(sMessage != null && sMessage.length > 0 ){
				alert(sMessage);
			}
			
			elem.focus();
			return false;
		} else {
			return true;
		}
	}	
	return true;
}

//************************************************************
// 目的：window.showModalDialog的方式開啟新的全螢幕視窗
// 傳回：
//************************************************************
function openFullScreen(theURL,winName) {
	var features="dialogWidth:"+window.screen.width+"px;dialogHeight:" + window.screen.height +"px;center=yes;help=no;status=no;resizable=no;scrollbars=yes";
	window.showModalDialog(theURL,winName,features);
}


//************************************************************
// 目的：ShowModalDialog
//	通常使用window.open的方式開啟新視窗的話
//	要取得父視窗的控制項，可以用window.opener來取得父視窗
//	然而如果使用showModalDialog的話...卻無效
//	如果有需要的話，需要修改開啟的語法以及showModalDialog中的語法
//	開啟語法第2個參數請下self,範例如下
//	var rc=window.showModalDialog(strURL,self,sFeatures);
//	然後接著就是呼叫父視窗的語法
//	var pWindow=window.dialogArguments;
//	這樣就可以取得父視窗的window物件控制了
// 傳回：
//************************************************************
function openShowModalDialog(sURL,w,h,obj){
    var sPath=sURL;
    if (obj.value!=undefined){
        sPath = sURL +obj.value ;
    } 	
	var dialogAnswer =window.showModalDialog(sPath, self,"dialogWidth:"+w+"px; dialogHeight:"+h+"px; center:yes;scroll:1;status:0;help:0;resizable:0");
	return dialogAnswer;
}

//************************************************************
// 目的：當USER按下IE右上角的x鈕時 開始處理.......
// 按下Button後Disable畫面上的所有Button PostBack回來後才恢復
//************************************************************
window.onbeforeunload=HandleOnClose;
function HandleOnClose(){
	//add by ted 20070110
	for (var j=0;j<document.all.length;j++) {

		if (document.all(j).tagName=="A"){
			document.all(j).disabled = true;
			document.all(j).removeAttribute("href");
		}
		if (document.all(j).tagName=="INPUT"){
			//還是要判斷一下 因為像type=text如果disabled 則換頁存檔時ViewState會不見
			if (document.all(j).type=="submit" || document.all(j).type=="button" || document.all(j).type=="reset"){
				if (document.all(j).disabled==false){
					document.all(j).disabled=true;
				}
			}
		}
	}

}


//************************************************************
// 目的：日期選擇
// 傳回：date string (eg. 2001/1/1)
//************************************************************
function openCalendar(sURL) {
	var sDate = window.showModalDialog(sURL,"calendar","dialogWidth:300px;dialogHeight:250px;center:1;scroll:0;help:0;status:0;resizable=0");
	sDate = (sDate == null) ? "" : sDate;
	return sDate;
}

//************************************************************
// 目的：自動調整 Control 的高度，預設5列
// 傳回：textareaHeight
//************************************************************
function textareaHeight(obj){
	if (obj.scrollHeight<=75){
		obj.style.posHeight= 75;
	} else {
		obj.style.posHeight=obj.scrollHeight;
	}
}

//************************************************************
// 字串取代函式
// replace(來源字串:String, 原始字串:String, 目標字串:String[, 取代數目:Number[, 是否逆向:Boolean]])
// 回傳值：取代結果字串
//************************************************************
String.prototype.replaceStr= function (str, a, b, n, bBeverse) {
	var counter = 0;
	var index = !bBeverse ? -1 : str.length;
	if (!isNaN(n)) {
		while (counter<=n) {
			index = !bBeverse ? str.indexOf(a, index+1) : str.lastIndexOf(a, index-1);
			if (index>-1) {
				counter++;
			} else {
				return str.split(a).join(b);
			}
		}
		return !bBeverse ? str.slice(0, index).split(a).join(b) +
		str.slice(index, str.length) : str.slice(0, index+1) +
		str.slice(index+1, str.length).split(a).join(b);
	} else {
		return str.split(a).join(b);
	}
}

//************************************************************
// 目的：中英文夾雜的字串長度：中文算2英文算1
// 傳回：
//************************************************************
function strLen(strParam) {
	if (strParam.length==0){
		return 0;
	}
	var intLen = strParam.length;
	var intTotal=0;
	for (i=0;i<intLen;i++) {
		intcode = strParam.charCodeAt(i);
		if (intcode > 255 || intcode < 0){
			intTotal = intTotal + 2
		}else{
			intTotal = intTotal + 1
		}
	}
	return intTotal;
}

/*************************************************************
目的    : 正確截取單字節和雙字節混和字符串
len		: 截取長度
回	:
*************************************************************/
 function substr(str, len) {
 	if(!str || !len) { 
 		return ''; 
 	}
	//預期計數：中文2字節，英文1字節     var a = 0;
	//循環計數
	var i = 0;
	//臨時字串
	var temp = '';
	for (i=0;i<str.length;i++){
		if (str.charCodeAt(i)>255){
			//按照預期計數增加2
			a+=2;
		}else{
			a++;
		}
		//如果增加計數後長度大於限定長度，就直接返回臨時字符串
		if(a > len) {
			return temp;
		}
		//將當前內容加到臨時字符串
		temp += str.charAt(i);
	}
	//如果全部是單字節字符，就直接返回源字符串
	return str;
}

//************************************************************
// 目的：檢查是否為英數字串
// 傳回：
//************************************************************
function isEnNum(obj) {
	var str =obj.value;
	//debug(str);
	// obj.value=obj.value.replace(/[^\u4E00-\u9FA50-9a-zA-Z_]/g,'');
	var re = /\W/; ///[^0-9a-zA-Z]/;
	if (re.test(str)) {
		obj.value="";
		//alert("只允許輸入英文及數字");
		return false;
	}
}

//************************************************************
// 目的：
// 傳回：
//************************************************************
function debug (text) {
	var div = document.createElement("div");
	div.appendChild(document.createTextNode(text));
	document.body.appendChild(div);
}

//************************************************************
// 目的：身份證字號驗證 and 公司統編驗證
// 傳回：
//************************************************************
function checkBothID (thisObj) {
	var id=thisObj.value;
	
	if (id.substr(0,3).toUpperCase() == 'OBU') {
		//alert('【OBU 編號】，請自行驗證授信戶統編！');
		return true;
	} else if (id.length == 10) {
		checkTWID( thisObj );
	} else if (id.length == 8) {
		checkCompanyID(thisObj);
	} else {
	    if (id.length > 0) {
		    alert('非八碼【統一編號】或 非十碼【身份證字號】，請自行驗證授信戶統編！');
		}    
	}	
}

//************************************************************
// 目的：身份證字號驗證
// 傳回：
//************************************************************
function checkTWID (thisObj) {
	var id=thisObj.value.toUpperCase();
	thisObj.value=id;
	var tab = "ABCDEFGHJKLMNPQRSTUVWXYZIO";
	var A1 = new Array (1,1,1,1,1,1,1,1,1,1,2,2,2,2,2,2,2,2,2,2,3,3,3,3,3,3 );
	var A2 = new Array (0,1,2,3,4,5,6,7,8,9,0,1,2,3,4,5,6,7,8,9,2,0,1,3,4,5 );
	var Mx = new Array (9,8,7,6,5,4,3,2,1,1);

	if ( id.length != 10 ) {
		return false;
	}
	var i = tab.indexOf( id.charAt(0) );
	if ( i == -1 ) {
		alert('【身份證字號】第一碼須為英文字母！');
		return false;
	}
	var sum = A1[i] + A2[i]*9;

	for ( i=1; i<10; i++ ) {
		var v = parseInt( id.charAt(i) );
		if ( isNaN(v) ) {
			alert('【身份證字號】第二碼到第十碼須為數字！');
			return false;
		}
		sum = sum + v * Mx[i];
	}
	if ( sum % 10 != 0 ) {
		alert('【身份證字號】編碼不正確，請重新輸入！');
		return false;
	}
	return true;
}


//************************************************************
// 目的：公司統編驗證
// 傳回：
//************************************************************
function checkCompanyID(thisObj){
	comNo=thisObj.value;
	if (!isNumber(comNo)){
		alert('請輸入【統一編號】數字八位');
		return false;
	}
	var res = new Array(8);
	var key = "12121241";
	var isModeTwo = false;	//第七個數是否為七
	var result = 0;

	if(comNo.length != 8){
		alert('【統一編號】不可少於或多於8碼！');
		return false ;
	}
	for(var i=0; i<8; i++){
		var tmp = comNo.charAt(i) * key.charAt(i);
		res[i] = Math.floor(tmp/10) + (tmp%10); //取出十位數和個位數相加
		if(i == 6 && comNo.charAt(i) == 7){
			isModeTwo = true;
		}
	}
	for(var s=0; s<8; s++){
		result += res[s];
	}
	if(isModeTwo){
		if((result % 10) != 0 && ((result + 1) % 10) != 0){//如果第七位數為7
			alert('【統一編號】編碼不正確，請重新輸入！');
			return false ;
		}
	}else if((result % 10) != 0){
		alert('【統一編號】編碼不正確，請重新輸入！');
		return false ;
	}
	return true;
}

//************************************************************
// 目的：檢查居留證號碼
// 傳回：
//
//	附件一：統一證號檢核原則
//一、統一證號編列規則：
//	共計十碼，第一碼為區域碼(同國民身分證)、第二碼為性別碼(入出境管理局使用ＡＢ；警察局外事科/課使用ＣＤ)、
//	第三至九碼為流水號、第十碼為檢查號碼。
//二、檢查號碼計算規則：
//	第一碼英文字母轉換為二位數字碼(轉換之數字與國民身分證同)，分別乘以特定數；第二碼英文字母轉換成二位數字後，
//	只取尾數乘以特定數；餘第三～九碼，亦分別乘以特定數。檢查號碼＝10－相乘後個位數相加總和之尾數。
//	惟若相乘後個位數相加總和尾數為0，則逕以「0」為檢查號碼。
//	舉例：FA12345689
//	(Ｆ：轉換為15，Ａ轉換為10─＞取尾數「0」)
//	【第一碼區域及第二碼性別之英文碼，先依據下列數字表換算，惟性別轉換後之二位數字碼，只取尾數。】
//	
//	A		B		C		D		E		F		G		H		J		K		L		M		N		P
//	10		11		12		13		14		15		16		17		18		19		20		21		22		23
//	Q		R		S		T		U		V		X		Y		W		Z		I		O				
//	24		25		26		27		28		29		30		31		32		33		34		35				
//	
//	     1501234568(統  號)
//	    ×1987654321(特定數)
//	     1507256528(不進位)
//	
//	     1＋5＋0＋7＋2＋5＋6＋5＋2＋8
//	     ＝41(將相乘後個位數相加)
//	    「41」(取尾數1───若尾數為0，則逕以「0」為檢查號碼)
//       檢查號碼＝10－1＝9
//三、基資登錄標準：
//    依據機器可判讀護照(Machine Readable Passport，簡稱ＭＲＰ護照)之編列規則登錄個人基本資料
//	  (先姓後名，姓名及護照號碼均不准登錄標點符號)。
//四、新舊居留證號轉碼方式：
//    例：舊號：A123456   ──>新號：AC01234567
//    說明：第一碼維持不變；第二碼依實際性別轉換為C或D；第三碼補０；
//	  第四至九碼帶入舊號之六位數流水號碼；第十碼依據前述檢查號碼計算規則計算得出。
//
//************************************************************
function residenceID(sender,e){
	var id=document.all("txt_id").value.toUpperCase();
	var kind=document.all("ddl_idkind").options[document.all("ddl_idkind").selectedIndex].text;

	if (id.length != 10) {
		return(e.IsValid=false);
	}
	
	if (isNaN(id.substr(2,8)) || (id.substr(0,1)<"A" ||id.substr(0,1)>"Z") || (id.substr(1,1)<"A" ||id.substr(1,1)>"Z")){
		return(e.IsValid=false);
	}

	var head="ABCDEFGHJKLMNPQRSTUVXYWZIO";
	id = (head.indexOf(id.substr(0,1))+10) +''+ ((head.indexOf(id.substr(1,1))+10)%10) +''+ id.substr(2,8);
	s =parseInt(id.substr(0,1)) +
	parseInt(id.substr(1,1)) * 9 +
	parseInt(id.substr(2,1)) * 8 +
	parseInt(id.substr(3,1)) * 7 +
	parseInt(id.substr(4,1)) * 6 +
	parseInt(id.substr(5,1)) * 5 +
	parseInt(id.substr(6,1)) * 4 +
	parseInt(id.substr(7,1)) * 3 +
	parseInt(id.substr(8,1)) * 2 +
	parseInt(id.substr(9,1)) +
	parseInt(id.substr(10,1));

	//判斷是否可整除
	if ((s % 10) != 0) {
		return(e.IsValid=false);
	}
	//居留證號碼正確
	return(e.IsValid=true);
}


 
//************************************************************
// 目的：將數字金額進行千位分隔
// 傳回：
//************************************************************
function formatNum(value){ 
	var objValue = stripCommas(value);	
	var digit = objValue.indexOf("."); // 取得小數點的位置 
	var int = objValue.substr(0,digit); // 取得小數中的整數部分 
	var i; 
	var mag = new Array(); 
	var word; 
	if (objValue.indexOf(".") == -1) { // 整數時 
		i = objValue.length; // 整數的個數 		

		while(i > 0) { 
			word = objValue.substring(i,i-3); // 每隔3位截取一組數字 			
			i-= 3; 
			mag.unshift(word); // 分別將截取的數字壓入數組 
		} 		 
		
		return mag; 
	}else{ // 小數時 
		i = int.length; // 除小數外，整數部分的個數 
		while(i > 0) { 
			word = int.substring(i,i-3); // 每隔3位截取一組數字 
			i-= 3; 
			mag.unshift(word); 
		} 
		return mag + objValue.substring(digit); 
	} 
}  

//************************************************************
// 目的：檢查MAIL格式
// 傳回：
//************************************************************
function isEMailAddr(thisObj) {
    var str = thisObj.value;
    if (str=="") { return true;}
    var re = /^[\w-]+(\.[\w-]+)*@([\w-]+\.)+[a-zA-Z]{2,7}$/;
    if (!str.match(re)) {
		alert("錯誤的 e-mail 格式!!");
        //setTimeout("focusElement('" + intRange.form.name + "', '" + intRange.name + "')", 0);
        return false;
    } else {
        return true;
    }
} 

//************************************************************
// 目的：強制focus
// 傳回：
//************************************************************
function focusElement(formName, elemName) {
    var elem = document.forms[formName].elements[elemName];
    elem.focus( );
    elem.select( );
}

//************************************************************
// 目的：
// 傳回：Object
//************************************************************
function getObjectByName(elementName) {
	var valid = false;

	if (elementName !=null){
		var element=window.document.getElementsByName(elementName);
		return element;
	}
	return null;
}

//************************************************************
// 目的：
// 傳回：Object
//************************************************************
function getObjectById(elementId) {
	var valid = false;
	alert(elementId);
	if (elementId !=null){	
		var element=window.document.getElementById(elementId);
		alert(element);
		return element;
	}
	
	return null;
}

//************************************************************
// 目的：檢查是否為數字 then 判斷位數,如果不符將textbox內容清空
//		intLen-->整數個數,floatLen-->小數點個數
// 傳回：
//************************************************************
function checkfloatLen(obj, intLen, floatLen){
	var restr =""
	var sValue=obj.value;
	sValue=sValue.replace(",", "");

	if(isNumber(sValue)){
		if (sValue.split(".")[1]==null){
			if (sValue.split(".")[0].length<=intLen){
				restr=sValue;
			}
		}else{
			if (sValue.split(".")[0].length<=intLen & sValue.split(".")[1].length<=floatLen){
				//this.document.all(obj.id).value=obj.value;
				restr=sValue;
			}
		}		
	}	
	obj.value=restr; 	
}

//************************************************************
// 目的：將數字金額進行千位分隔","
// 傳回：
//************************************************************
function formatNumPrint(obj){	
	var arr =formatNum(obj.value);
	var str="";
	for(var i=0; i<arr.length; i++){
		str=str+","+arr[i];
	}
	obj.value=str.substring(1,str.length); 
} 

//************************************************************
//'[帳號檢核]
//
//    '帳號檢核規則
//    '帳號:  9  8  3  0  0  1  0  6  0  3  3  4
//    '	  ＊  6  5  4  3  2  7  6  5  4  3  2
//    '9*6+8*5+3*4+0*3+0*3+1*7+0*6+6*5+0*4+3*3+3*2 = 158
//
//    '158/9 →商數= 17       餘數=5     檢查碼=9-5→ 4
//
//************************************************************
function checkAccount(thisObj){
	AccountNo=thisObj.value.trim();
	if (!isNumber(AccountNo)){
		alert('請輸入【帳號】數字12位');
		return false;
	}
	
	var key = "65432765432";
	var sum = 0;

	if(AccountNo.length != 12){
		alert('【帳號】不可少於或多於12碼！');
		return false ;
	}
	
	for(var i=0; i<11; i++){
		var tmp = AccountNo.charAt(i) * key.charAt(i);
		sum = sum+tmp;
	}
 	
 	sum=sum%9;
 	
 	if((9-sum)!=AccountNo.charAt(11)){
 		alert('【帳號】編碼不正確，請重新輸入！');
		return false ;
 	}
	thisObj.value=AccountNo;
	return true;

}

//************************************************************
// 目的：將字串轉換成數字
// 傳回：
//************************************************************
function shgNumber(strNum){

	if (isEmptyStr(strNum)){
		return false;
	}

	var num = new Number(stripCommas(strNum));
	if (isNaN(num)){
		return false;
	}
	return num;
}

function Fill7Zero() {
    var vActNo = document.all["txtActNO"].value;
    var chkNum;
    var vActSEQ = document.all["txtActSEQ"].value;

    if (vActNo.length == 12) {
        chkNum = vActNo.substr(3, 2)
        if (chkNum == '41') {
            document.all["txtActSEQ"].value = padLeft(vActSEQ, 7);
        }
    }
}

function Fill7ZeroByClientID(txtActNO, txtActSEQ) {
    var vActNo = document.all(txtActNO).value;
    if (vActNo.length == 0) {
        alert("請輸入帳號!");
        document.all(txtActNO).select();
        document.all(txtActNO).focus();
        return false;
    }

    var chkNum;

    var vActSEQ = document.all(txtActSEQ).value;
    if (vActSEQ.length == 0) {
        alert("請輸入分號!");
        document.all(txtActSEQ).select();
        document.all(txtActSEQ).focus();
        return false;
    }

    if (vActNo.length == 12) {
        chkNum = vActNo.substr(3, 2)
        if (chkNum == '41') {
            if (isNaN(vActSEQ)) {
                alert("分號應為數字!請檢查!");
                document.all(txtActSEQ).select();
                document.all(txtActSEQ).focus();
                return false;                
            }
            else {
                document.all(txtActSEQ).value = padLeft(vActSEQ, 7);
            }            
        }
        else {
            if (vActSEQ.length!=7)
            {
                alert("分號應為7位數字!請檢查!");
                document.all(txtActSEQ).select();
                document.all(txtActSEQ).focus();
                return false;
            }

            if (isNaN(vActSEQ)) {
                alert("分號應為7位數字!請檢查!");
                document.all(txtActSEQ).select();
                document.all(txtActSEQ).focus();
                return false;
            }
        }
    }
    else {
        alert("帳號應為12碼!請檢查!");
        document.all(txtActNO).select();
        document.all(txtActNO).focus();
        return false;
    }

    return true;
}

//左邊補零
function padLeft(str, lenght) {
    if (str.length >= lenght)
    { return str; }
    else
    { return padLeft("0" + str, lenght); }
}
 