<!--
// <SCRIPT language=javascript type=text/javascript>
var inGridName;
//function KCParseDec(s)    { return( parseFloat(KCTrim(s.replace(/,/g,"")))); }
function KCLTrim(s){ 
	return( ("string"!=typeof(s))? null : s.replace(/^ +/,"").replace(/^ +/,"").replace(/^ +/,"") ); 
}
function KCRTrim(s){ 
	return( ("string"!=typeof(s))? null : s.replace(/ +$/,"").replace(/ +$/,"").replace(/ +$/,"") ); 
}
function KCTrim(s){
	return( KCRTrim( KCLTrim(s) ) );
}

//modify by ted 20060815
function KCParseDec(s){
    if (isNaN(parseFloat(KCTrim(s.replace(/,/g,""))))){
		return 0;
	}else{
		return (parseFloat(KCTrim(s.replace(/,/g,""))));
	}
}

var KC_on_shift;
var KC_BACKSPACE    =8;
var KC_TAB          =9;
var KC_RETURN       =13;
var KC_SHIFT        =16;
var KC_CTRL         =17;
var KC_ALT          =18;
var KC_PAUSE        =19;
var KC_CAPSLOCK     =20;
var KC_NUMLOCK      =144;
var KC_SCROLLLOCK   =145;
var KC_PRTSCM       =106;
var KC_ESCAPE       =27;
var KC_SPACE        =32;
var KC_PAGEUP       =33;
var KC_PAGEDOWN     =34;
var KC_END          =35;
var KC_HOME         =36;
var KC_LEFT         =37;
var KC_RIGHT        =39;
var KC_UP           =38;
var KC_DOWN         =40;
var KC_INSERT       =45;
var KC_DELETE       =46;
var KC_F1           =112;
var KC_F2           =113;
var KC_F3           =114;
var KC_F4           =115;
var KC_F5           =116;
var KC_F6           =117;
var KC_F7           =118;
var KC_F8           =119;
var KC_F9           =120;
var KC_F10          =121;
var KC_F11          =122;
var KC_F12          =123;
var KC_UF1          =234;//EA
var KC_UF2          =249;//F9
var KC_UF3          =245;//F5
var KC_UF4          =243;//F3
var KC_UF5          =126;//7E
var KC_UF6          =127;//7F+sendkey(13)
var KC_UF7          =251;//FB
var KC_UF8          = 47;//2F
var KC_UF9          =124;//7C
var KC_UF10         =125;//7D
var KC_PAD_0        = 96;
var KC_PAD_1        = 97;
var KC_PAD_2        = 98;
var KC_PAD_3        = 99;
var KC_PAD_4        =100;
var KC_PAD_5        =101;
var KC_PAD_6        =102;
var KC_PAD_7        =103;
var KC_PAD_8        =104;
var KC_PAD_9        =105;
var KC_PAD_MUL      =106;
var KC_PAD_ADD      =107;
var KC_PAD_SUB      =109;
var KC_PAD_DOT      =110;
var KC_PAD_DIV      =111;
var KC_COMMA        =188;
var KC_MINUS        =189;
var KC_PERIOD       =190;
var KC_SLASH        =191;

var errItem = "";
var CAPITALS  = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
var rtnValue;

function KCAmtCheck(){
//  alert ("keycode=" + window.event.keyCode + ", shiftKey=" + window.event.shiftKey);
	if (!window.event.shiftKey) {
		var c = String.fromCharCode(event.keyCode);
		if (window.event.keyCode== KC_BACKSPACE )       {return true;}
		if (window.event.keyCode== KC_TAB       )       {return true;}
		if (window.event.keyCode== KC_END       )       {return true;}
		if (window.event.keyCode== KC_HOME      )       {return true;}
		if (window.event.keyCode== KC_LEFT      )       {return true;}
		if (window.event.keyCode== KC_RIGHT     )       {return true;}
		if (window.event.keyCode== KC_INSERT    )       {return true;}
		if (window.event.keyCode== KC_DELETE    )       {return true;}
		if (window.event.keyCode== KC_NUMLOCK   )       {return true;}
		if (window.event.keyCode== KC_CAPSLOCK  )       {return true;}
		if (window.event.keyCode== KC_COMMA     )       {return true;}
		if (window.event.keyCode== KC_MINUS     )       {return true;}
		if (window.event.keyCode== KC_PERIOD    )       {return true;}
		if (window.event.keyCode== KC_PAD_0     )       {return true;}
		if (window.event.keyCode== KC_PAD_1     )       {return true;}
		if (window.event.keyCode== KC_PAD_2     )       {return true;}
		if (window.event.keyCode== KC_PAD_3     )       {return true;}
		if (window.event.keyCode== KC_PAD_4     )       {return true;}
		if (window.event.keyCode== KC_PAD_5     )       {return true;}
		if (window.event.keyCode== KC_PAD_6     )       {return true;}
		if (window.event.keyCode== KC_PAD_7     )       {return true;}
		if (window.event.keyCode== KC_PAD_8     )       {return true;}
		if (window.event.keyCode== KC_PAD_9     )       {return true;}
		if (window.event.keyCode== KC_PAD_DOT   )       {return true;}
		if (window.event.keyCode== KC_PAD_ADD   )       {return true;}
		if (window.event.keyCode== KC_PAD_SUB   )       {return true;}
		if ("0123456789.,-+".indexOf(c,0) >= 0  )       {return true;}
    }
    return false;
}

function KCNumCheck(){
  //alert ("keycode=" + window.event.keyCode + ", shiftKey=" + window.event.shiftKey);
	if (!window.event.shiftKey) {
		var c = String.fromCharCode(event.keyCode);
		if (window.event.keyCode		== KC_BACKSPACE )       {return true;}
		if (window.event.keyCode		== KC_TAB       )       {return true;}
		if (window.event.keyCode		== KC_END       )       {return true;}
		if (window.event.keyCode		== KC_HOME      )       {return true;}
		if (window.event.keyCode		== KC_LEFT      )       {return true;}
		if (window.event.keyCode		== KC_RIGHT     )       {return true;}
		if (window.event.keyCode		== KC_INSERT    )       {return true;}
		if (window.event.keyCode		== KC_DELETE    )       {return true;}
		if (window.event.keyCode		== KC_NUMLOCK   )       {return true;}
		if (window.event.keyCode		== KC_CAPSLOCK  )       {return true;}
		if (window.event.keyCode		== KC_PAD_0     )       {return true;}
		if (window.event.keyCode		== KC_PAD_1     )       {return true;}
		if (window.event.keyCode		== KC_PAD_2     )       {return true;}
		if (window.event.keyCode		== KC_PAD_3     )       {return true;}
		if (window.event.keyCode		== KC_PAD_4     )       {return true;}
		if (window.event.keyCode		== KC_PAD_5     )       {return true;}
		if (window.event.keyCode		== KC_PAD_6     )       {return true;}
		if (window.event.keyCode		== KC_PAD_7     )       {return true;}
		if (window.event.keyCode		== KC_PAD_8     )       {return true;}
		if (window.event.keyCode		== KC_PAD_9     )       {return true;}
		if ("0123456789".indexOf(c,0) 	>= 0			)       {return true;}
	}
	return false;
}

function KCRocDateCheck(thisObj){
	if(thisObj==null){
		thisObj = window.event.srcElement;
	}
	thisObj.value = KCTrim(thisObj.value);
	//0305add
	//-------
	if (thisObj.value.length==0){
		return true;
	}
	var arrayYMD = thisObj.value.split(".")
	if (arrayYMD.length == 3) {
		thisObj.value = arrayYMD[0] + ("00"+arrayYMD[1]).substr(("00"+arrayYMD[1]).length-2,2) + ("00"+arrayYMD[2]).substr(("00"+arrayYMD[2]).length-2,2)
	}
	if (thisObj.value.length < 5){
		thisObj.select();
		thisObj.focus();
		alert ("日期格式錯誤『YYYMMDD』 ");
		return false;
	}

	if (KCParseDec(thisObj.value)==NaN) {
		thisObj.select();
		thisObj.focus();
		alert ("日期格式錯誤『YYYMMDD』 ");
		return false;
	}
	var YYYYMMDD = KCParseDec(thisObj.value) + 19110000;
	var DD    =  YYYYMMDD % 100;
	var MM    = (YYYYMMDD % 10000 - DD) / 100 ;
	var YYYY  = (YYYYMMDD - (MM) * 100 - DD) / 10000;
	if (!KCDateCheck(YYYY,MM,DD)) {
		thisObj.select();
		thisObj.focus();
		alert ("日期錯誤『YYYMMDD』 ，請輸入正確日期");
		return false;
	}
	thisObj.value = KCParseDec(thisObj.value);
	if (arrayYMD.length == 3) {
		thisObj.value = arrayYMD.join(".");
	}
	if(thisObj.value.length == 6){
	  thisObj.value = "0"+thisObj.value;
	}
	if(thisObj.value.length == 5){
	  thisObj.value = "00"+thisObj.value;
	}
	return true;
}
//把日期轉為YYY.MM.DD型式
function FormatDate(obj) {
            
            //alert(str);			
            var str = obj.value;
            var tmpstr = str.trim().replaceStr(str, ".", "", 2);
            var len = tmpstr.length;
            if (len == 6) {
                if (tmpstr.substr(0, 1) != 0) {
                    if (isValidTWDate(str)) {
                        if (str != tmpstr.substr(0, 2) + "." + tmpstr.substr(2, 2) + "." + tmpstr.substr(4, 2)) {
                            obj.value = tmpstr.substr(0, 2) + "." + tmpstr.substr(2, 2) + "." + tmpstr.substr(4, 2);
                        }
                    } else {
                        obj.value = "";
                    }
                }
            } else if (len == 7) {
                if (tmpstr.substr(0, 1) != 0) {
                    if (isValidTWDate(str)) {
                        if (str != tmpstr.substr(0, 3) + "." + tmpstr.substr(3, 2) + "." + tmpstr.substr(5, 2)) {
                            obj.value = tmpstr.substr(0, 3) + "." + tmpstr.substr(3, 2) + "." + tmpstr.substr(5, 2);
                        }
                    } else {
                        obj.value = "";
                    }
                } else {
                    if (isValidTWDate(str)) {
                        if (str != tmpstr.substr(1, 2) + "." + tmpstr.substr(3, 2) + "." + tmpstr.substr(5, 2)) {
                            obj.value = tmpstr.substr(1, 2) + "." + tmpstr.substr(3, 2) + "." + tmpstr.substr(5, 2);
                        }
                    } else {
                        obj.value = "";
                    }
                }
            }
        };

        //檢核YYYMMDD日期 + 把日期轉為YYY.MM.DD型式
function Check_YYYMMDDDate(obj) {
        if (KCRocDateCheck(obj)) {
        FormatDate(obj);
        }
    };

function KCDateCheck(YYYY,MM,DD){
	if(isNaN(YYYY) || isNaN(MM) || isNaN(DD)){
		return false;	
	}
	var tDate = new Date(YYYY, MM-1, DD);
	var yyyy  = tDate.getYear();
	var mm    = tDate.getMonth()+1;
	var dd    = tDate.getDate();
	if ((YYYY%100)!=(yyyy%100) || MM!=mm || DD!=dd) {
		return false;
	}
	return true;
}

function KCMmCheck(){
	window.event.srcElement.value = KCTrim(window.event.srcElement.value);
	if (window.event.srcElement.value.length==0){
		return true;
	}
	if(KCParseDec(window.event.srcElement.value)==NaN) {
		window.event.srcElement.select();
		window.event.srcElement.focus();
		alert ("『月』 輸入錯誤，請重新輸入");
		return false;
	}
	var MM = KCParseDec(window.event.srcElement.value);
	if(MM<1 || MM>12) {
		window.event.srcElement.select();
		window.event.srcElement.focus();
		alert ("『月』 輸入錯誤，請重新輸入");
		return false;
	}
	return true;
}

function KCDdCheck(){
	window.event.srcElement.value = KCTrim(window.event.srcElement.value);
	if (window.event.srcElement.value.length==0){
		return true;
	}
	if(KCParseDec(window.event.srcElement.value)==NaN) {
		window.event.srcElement.select();
		window.event.srcElement.focus();
		alert ("『日』 輸入錯誤，請重新輸入");
		return false;
	}
	var MM = KCParseDec(window.event.srcElement.value);
	if(MM<1 || MM>31) {
		window.event.srcElement.select();
		window.event.srcElement.focus();
		alert ("『日』 輸入錯誤，請重新輸入");
		return false;
	}
	return true;
}

function KCAmtFormat(indata){
	window.event.srcElement.value = KCTrim(window.event.srcElement.value);
	if(window.event.srcElement.value.length==0 || KCParseDec(window.event.srcElement.value)==NaN){
		return false;
	}
	window.event.srcElement.value = KCAmtFmt(KCParseDec(window.event.srcElement.value), indata);
	return true;
}

function KCDotFormat(fraction){
	//alert(window.event.srcElement.value);
	window.event.srcElement.value = KCTrim(window.event.srcElement.value);
	if(window.event.srcElement.value.length==0 || KCParseDec(window.event.srcElement.value)==NaN){
		return false;
	}
	window.event.srcElement.value = KCAmtFmt(KCParseDec(window.event.srcElement.value), fraction);
	return true;
}


function KCChk100Percent(){
	var fraction = 2;
	if(KCChk100Percent.arguments.length > 0) {
		fraction = KCChk100Percent.arguments[0];
	}
	if(!KCDotFormat(fraction)){
		return false;
	}
	var iValue = KCParseDec(window.event.srcElement.value);
	if (iValue <= 100 && iValue >= -100){
		return true;
	}
	alert ("比率值輸入不合常理 ，請檢查 ");
	window.event.srcElement.select();
	window.event.srcElement.focus();
	return false;
}


function KCChkPercent100_4(){
	var fraction = 4;
	if(KCChkPercent100_4.arguments.length > 0) {
		fraction = KCChkPercent100_4.arguments[0];
	}
	if(!KCDotFormat(fraction)){
		return false;
	}
	var iValue = KCParseDec(window.event.srcElement.value);
	if (iValue <= 100 && iValue >= -100){
		return true;
	}
	alert ("比率值輸入不合常理 ，請檢查 ");
	window.event.srcElement.select();
	window.event.srcElement.focus();
	return false;
}


function KCCalFmtPercent(numerator, denominator, quotient, fraction){  // numerator / denominator
	var Vnumerator   = KCParseDec(numerator.value);
	var Vdenominator = KCParseDec(denominator.value);
	var Vquotient;
	quotient.value = "";
	if (Vdenominator==0){
		Vquotient = 0;
		quotient.value = KCAmtFmt(Vquotient, fraction);
	}
	if (!isNaN(Vnumerator) && !isNaN(Vdenominator) && Vdenominator!=0) {
		Vquotient = Vnumerator * 100 / Vdenominator;
		if (fraction > 0){
			var Vround = 0.5;
			for (var i=0; i< fraction; i++) {
				Vround /= 10;
			}
			Vquotient += Vround;
		}
		quotient.value = KCAmtFmt(Vquotient, fraction);
	}
	return true;
}

function KCCalFmtPercent1(numerator, denominator, quotient, fraction){ // (numerator - denominator) / denominator

	var Vnumerator   = KCParseDec(numerator.value);
	var Vdenominator = KCParseDec(denominator.value);
	var Vquotient;
	quotient.value = "";
	if (!isNaN(Vnumerator) && !isNaN(Vdenominator) && Vdenominator!=0) {
		Vquotient = (Vnumerator - Vdenominator) * 100 / Vdenominator;
		if (fraction > 0){
			var Vround = 0.5;
			for (var i=0; i< fraction; i++) {
				Vround /= 10;
			}
			Vquotient += Vround;
		}
		quotient.value = KCAmtFmt(Vquotient, fraction);
	}
	return true;
}

function KCAmtFmtSubtract(minuend, subtrahend, remainder, idx){
	var Vminuend    = KCParseDec(minuend.value);
	var Vsubtrahend = KCParseDec(subtrahend.value);
	var Vremainder  = KCParseDec(remainder.value);
	if (idx==1){
		if (!isNaN(Vsubtrahend) && !isNaN(Vremainder)) {
			Vminuend = Vsubtrahend + Vremainder;
			minuend.value = KCAmtFmt(Vminuend, 0)
			return true;
		}
	} else if (idx==2){
		if (!isNaN(Vminuend) && !isNaN(Vremainder)) {
			Vsubtrahend = Vminuend - Vremainder;
			subtrahend.value = KCAmtFmt(Vsubtrahend, 0)
			return true;
		}
	} else	if (idx==3){
		if (!isNaN(Vminuend) && !isNaN(Vsubtrahend)) {
			Vremainder = Vminuend - Vsubtrahend;
			remainder.value = KCAmtFmt(Vremainder, 0)
			return true;
		}
	}
	return false;
}

function KCAmtFmt(Vquotient, fraction){
	if (isNaN(fraction)){
			fraction=0;
	}
	
	if(isNaN(KCParseDec(String(Vquotient)))){
		return "";
	}
	var num = String(KCParseDec(String(Vquotient)));
	
	//alert("first num="+num)
	 var tmpNum;
  
	var ret,sign,dec,dot;
	ret = sign = dec = "";
	if( num.charAt(0) == '-' )	{
		sign = '-';
		num = num.substr(1);
		
		if (fraction>0)	{
			tmpnum = parseFloat(num).toFixed(fraction);
		}else{
			tmpnum = (num*1).toFixed(0);
		}
		
		num = String(tmpnum);
	}else if( num.charAt(0) == '+' )	{
		sign = '+';
		num = num.substr(1);
		
		if (fraction>0)		{
			tmpnum = parseFloat(num).toFixed(fraction);
		}else{
			tmpnum = (num*1).toFixed(0);
		}
		
		num = String(tmpnum);
	} else  {
  		sign = '';
  		num = num.substr(0);
  		//alert("num="+num);
	  	
  		if (fraction>0)	{
  			tmpnum = parseFloat(num).toFixed(fraction);
  		}else{
  			tmpnum = (num*1).toFixed(0);
  		}
	  	
  		//alert("tmpnum="+tmpnum);
  		num = String(tmpnum);
	}
  
	dot = num.indexOf('.');
	
	if( dot>= 0 ){
		dec = num.substr(dot,fraction+1);
		num = num.substr(0,dot);
		if(dec.length==1){
			dec = "";
		}
	}
	if ( dot < 0) {
		  dec = "";
	}
	if( num.length <= 3 ){
		if( num.length == 0 ) num = '0';
		ret = sign + num + dec;
	}else{
		var i = num.length % 3;
		if( i>0 ){
			ret = num.substr(0,i)+',';
		}
		while( (i+3) < num.length ){
			ret += num.substring(i,i+3)+',';
			i += 3;
		}
		ret += num.substr(i);
		ret = sign + ret + dec;
	}
	//alert("ret="+ret);
	return ret;
}

function KCmaxWord(text,max){
	if(text.value.length > max){
		text.value=text.value.substr(0,max);
	}
}

function KCQuotationCheck(){
	var sInput = KCTrim(window.event.srcElement.value);
	if (sInput.indexOf("'") >-1){
		sInput =sInput = window.event.srcElement.value.replace(/'/g,"");
		window.event.srcElement.value=sInput;
	}
}

function FormatNumber(){
  if(window.event.keyCode==16)  {
	shift_tmp = 0;
  }
  if (window.event.srcElement.value.charAt(0)=="-")  {
	a=1;
	window.event.srcElement.value=window.event.srcElement.value.replace(/,/g,"").substring(1,window.event.srcElement.value.replace(/,/g,"").length);
  } else  {  　	
	a=0;
  }
  var sInput = window.event.srcElement.value.replace(/,/g,"");
//不能打小數點  
//  if (sInput.length >=0)
//  var kdot = sInput.indexOf('.');
//  {
//  if (kdot > -1)
//  {
//  	sInput = sInput.substring(0,kdot);
//  }
//  else
//  {
//  	sInput = sInput;
//  }
//  window.event.srcElement.value = sInput;
//  }  
　if (sInput.length >= 3)　{
　　var sHead = sInput.substring(0,(sInput.length % 3));
　　var sTail = "";
　　for (var i = ((sInput.length) % 3) ; i < sInput.length ; )　{
　　　sTail = sTail + "," + sInput.substring(i,i+3);
　　　i+=3;
　　}
　　if (sHead == "")　{
　　　sTail = sTail.substring(1,sTail.length);
　　}
		
　　window.event.srcElement.value = (sHead + sTail);
　}
  if (a==1)  {
	window.event.srcElement.value="-"+window.event.srcElement.value;
  }
}


function bgIDNcheck(thisObj,si){
	//if (!Lock(thisObj)){
	// return false;
	//}
	thisObj.value=thisObj.value.toUpperCase();
	var tID = thisObj.value;
	var tlength = tID.length;
	//alert(tID);
	//alert(tlength);
	//alert(thisObj.id);
	for (var i=0; i<tlength; i++){
  		var chk = tID.charAt(i);
  		if (chk==" "){
  			alert("身份証統編輸入錯誤？");
  			return checkError(thisObj);
  		}
	}
	switch(si){
		case "A" :checkIDN(thisObj);break;
		case "B" :checkBAN(thisObj);break;
		case "C" : break;
		default  :alert("身份証統編輸入錯誤？");return checkError(thisObj);
	}
}

//檢查身份証字號
function checkIDN(thisObj){	
	//alert(Lock(thisObj));
	//if (!Lock(thisObj)) return false;
	thisObj.value=thisObj.value.toUpperCase();
	var myID=KCTrim(thisObj.value);
	//alert(ID);
	if (myID.length==0){
		return checkOK(thisObj);
	}
	if (myID.length!=10) {
		alert("身分證字號必須為１０碼？");
		return checkError(thisObj);
	}
	var aryid= new Array(10);
	for (var i=0; i<myID.length; i++) {
		aryid[i]=myID.charAt(i);
	}
	aryid[0]=CAPITALS.indexOf(aryid[0]);
	if(aryid[0]==-1) {
		alert('身份証字號第一碼不為英文字母？');
		return checkError(thisObj);
	}
	if (aryid[1]!=1 && aryid[1]!=2) {
		alert('身份証字號無法辦識性別？')
		return checkError(thisObj);
	}
	if (isNaN(myID.substring(2))) {
		alert('身份証字號後８碼不為數字？')
		return checkError(thisObj);
	}
	var code=new Array(26)
	code[ 0]= 1;code[ 1]=10;code[ 2]=19;code[ 3]=28;code[ 4]=37;code[ 5]=46;
	code[ 6]=55;code[ 7]=64;code[ 8]=39;code[ 9]=73;code[10]=82;code[11]= 2;
	code[12]=11;code[13]=20;code[14]=48;code[15]=29;code[16]=38;code[17]=47;
	code[18]=56;code[19]=65;code[20]=74;code[21]=83;code[22]=21;code[23]= 3;
	code[24]=12;code[25]=30;
	var result=code[aryid[0]]
	for (var i=1; i<myID.length; i++) {
		result+=aryid[i]*(9-i);
		
	}
	result+=1*aryid[9]
	if (result%10!=0) {
		alert("身份証字號加總檢核碼錯誤？");
		return checkError(thisObj);
	}
	return checkOK(thisObj);
}

function checkBAN(thisObj){
	//if (!Lock(thisObj)) return false;
	var myID=KCTrim(thisObj.value);
	//alert(myID);
	if (myID.length==0){
		return checkOK(thisObj);
	}
	if (myID.length!=8 || isNaN(myID.substring(4,4))) {
		alert("法人統一編號必須為８碼，且後４碼須為數字");
		return checkError(thisObj);
	}
	//var c1 = (ID.charAt(0) * 1);
	//var c2 = (ID.charAt(2) * 1);
	//var c3 = (ID.charAt(4) * 1);
	//var c4 = (ID.charAt(7) * 1);
	//var b1 = (ID.charAt(1) * 2)      % 10;
	//var a1 = (ID.charAt(1) * 2 - b1) / 10;
	//var b2 = (ID.charAt(3) * 2)      % 10;
	//var a2 = (ID.charAt(3) * 2 - b2) / 10;
	//var b3 = (ID.charAt(5) * 2)      % 10;
	//var a3 = (ID.charAt(5) * 2 - b3) / 10;
	//var b4 = (ID.charAt(6) * 4)      % 10;
	//var a4 = (ID.charAt(6) * 4 - b4) / 10;
	//var b5 = (a4 + b4)      % 10;
	//var a5 = (a4 + b4 - b5) / 10;
	//var Y = a1 + b1 + c1 + a2 + b2 + c2 + a3 + b3 + c3 + a4 + b4 + c4;
	//if ((Y % 10) == 0)
	//  return checkOK(thisObj);
	//if (ID.charAt(7) == 7) {
	//  Y = a1 + b1 + c1 + a2 + b2 + c2 + a3 + b3 + c3 + a4 + a5 + c4;
	//  if ((Y % 10) == 0)
	//      return checkOK(thisObj);
	//}
	//alert('統一編號加總檢核碼錯誤？');
	//return checkError(thisObj);
	var the_ban = new Array();
	var i, j , sum ;
	var weight = new Array( 1, 2, 1, 2, 1, 2, 4, 1);
	sum = 0 ;
	for ( i = 0 ; i < 8 ; i ++ ) {
		the_ban[i] = myID.charAt(i) ;
		if ( isNaN(the_ban[i]) == true ) {
			if ( i < 4 ) {
				j = myID.charCodeAt(i) - 48 ;
				// equivelenat to 'A' - '0' in C language
				// 48 is the ascii code of '0' in decimal
				j = ( ( myID.charCodeAt(i) - 48 ) % 10 ) * weight[i];
			}else {
				j = i+1;
				alert("法人統一編號應為８碼數字,虛擬統編應為４碼字母４碼數字, 但輸入值的第 "+j+" 碼不為數字");
				return checkError(thisObj);
			}
		}else {
			j = eval(the_ban[i]) * weight[i] ;
		}
		sum += j % 10 ;
		sum += Math.floor( j/10 );
	}
	sum = sum % 10;
	if ( sum == 0 || ( sum == 9 && the_ban[6] == '7') ) {
		return checkOK(thisObj);
	}else {
		alert('統一編號加總檢核碼錯誤？');
	}
	return checkError(thisObj);
}

function Lock(thisObj){
	if (errItem!="" && errItem!=thisObj.id){
		return false;
	}
	return true;
}
function checkError(thisObj){
	thisObj.focus();
	thisObj.select();
	errItem=thisObj.id;
	return false;
}
function checkOK(thisObj){
	errItem="";
	return true;
}

function checkLen(oSrc, nLen, sText){
	var bEvent = true;
	if (event.type == "keypress")	{
		if (event.keyCode < 32){
			return true;
		}
		// 取得text內容
		sText = getKeypressText();
		// 檢查輸入字元個數
		if  (sText.length + sText.replace(/[ -~]/g, "").length > nLen)		{
			event.returnValue = false;
			return false;
		}else{
			return true;
		}
	}else if (event.type == "blur")	{
		var sRtnText = "";
		for (var i=1;i<=sText.length;i++){
			var sTemp = sText.substr(0,i);
			if (sTemp.length + sTemp.replace(/[ -~]/g, "").length <= nLen){
				sRtnText = sTemp;
			}
		}
		oSrc.value = sRtnText;
	}
}

function getKeypressText(){
	var oRange = document.selection.createRange();
	oRange.text = String.fromCharCode(event.keyCode);

	var sText = event.srcElement.value;
	oRange.moveStart("character", -1);
	oRange.text = "";
	return sText;
}

//防止User type -> ' (單引號), Event -> onkeypress
//document.onkeypress = handlerkeypress;
function handlerkeypress(text){
	//alert(window.event.keyCode);
	//alert("no way!")
	if (event.keyCode == 222 ) { 
		text.value = text.value.substr(0, text.value.length -1); 
	}
	if (event.keyCode ==  39) { 
		event.returnValue = false;  
	}
}

function maxWord(text,max)
{
	 var addascii=0 ;
	 for(var i=0;i<text.value.length  ;i++)	 {
		var codes=text.value.charCodeAt(i);
		//alert(codes);
		if(codes>=32 && codes<=10000)		{
			addascii++;
		}else{
			addascii+=2;
		}

		if(addascii > max){
			text.value=text.value.substr(0,i);
			alert( "已輸入超過"+max+"字元長度");
			return false;
		}
	}
}

/**
* Validate the input text as a number
*
* @param object ID of the input control object
* @param string Number type ('int', 'u_int', 'float' or 'u_float')
* @param number Max number allowed to input
*
* @return null
*/
/*使用方法 sample:DB欄位為Number(7,2)則不可大於等於100000
onkeypress="javascript:KCAmtCheck();" onchange="validationNumber(this, 'float', 100000);KCDotFormat(2);"
*/
function validationNumber(hndID, numType, maxNum){
	var keyCode = window.event.keyCode;
	var ch = String.fromCharCode(keyCode);
	var value = '';
	var retCode = 0;
	var oNumVerify=null;
	//擋Enter鍵
	if (keyCode==13){
		window.event.keyCode = 0;
		return null;
	}
	if (hndID != undefined)	{
		value += hndID.value+ch;
	}else{
		window.event.keyCode = keyCode;
		return null;
	}
	value = KCParseDec(value);
	oNumVerify = new isNumeric(value);
	if (oNumVerify.isNumber){
		numType = numType.toString().toLowerCase();
		switch(numType)	{
			case 'u_int':  //正整數
			if ((oNumVerify.isMinus==false) && (oNumVerify.isDecimal==false)){
			retCode = keyCode;
			}
			break;
		case 'u_float':  //正實數
			if (oNumVerify.isMinus==false){
				retCode = keyCode;
			}
			break;
		case 'int':   //整數
			if (oNumVerify.isDecimal==false){
			retCode = keyCode;
			}
			break;
		case 'float':  //實數
			retCode = keyCode;
			break;
		default:
			retCode = 0;
		}
		//判斷輸入的數字是否超過設置的最大值
		if ((maxNum!=undefined) && (maxNum.constructor==Number) && (oNumVerify.value >= maxNum))	{
			retCode = 0;
			alert("值不能大於等於"+maxNum);
			window.event.srcElement.select();
			window.event.srcElement.focus();
			/*
			if ((hndID != undefined) && (hndID.type="text") )
			{
	 			hndID.value="";
			}
			*/
		}
	}
	oNumVerify = null;
	window.event.keyCode = retCode; 
	return null;
} 

function isNumeric(verifyNum){
	var re = /^([-]{0,1})([0-9]*)([\.]{0,1})([0-9]*)$/g;
	this.isNumber = false;
	this.isMinus = false;
	this.isDecimal = false;
	this.value = verifyNum;
	verifyNum = verifyNum.toString();
	if (re.test(verifyNum))	{
		this.isNumber = true;
		re.exec(verifyNum);
		//判斷 '-' 符號
		if (RegExp.$1=='-')	{
			this.isMinus = true;
		}
		//判斷 '.' 符號
		if (RegExp.$3=='.')	{
			this.isDecimal = true;
			verifyNum += '0';
		}
		try	{
			this.value = parseFloat(verifyNum);
		}catch(e){
			this.value = 0;
		}
	}
	return;
}


//防止User type -> ' (單引號), Event -> onkeypress
/*
function KCWordCheck()
{
	//if ((event.keyCode >= 33 && event.keyCode <= 39) || (event.keyCode >= 58 && event.keyCode <= 64) ||(event.keyCode >= 91 && event.keyCode <= 96) || (event.keyCode >= 123 && event.keyCode <= 126) || event.keyCode == 44)
	//	event.returnValue = false;
	
	if (event.keyCode ==  39) {
		event.returnValue = false;
		}
}
*/

//************************************************************
// 目的：當USER按下IE右上角的x鈕時 開始處理.......
// 按下Button後Disable畫面上的所有Button PostBack回來後才恢復
//************************************************************
window.onbeforeunload=HandleOnClose;
function HandleOnClose()
{
	if (event.clientY < 0) {
		//if(window.confirm('確定離開嗎?')!=true){
		//event.returnValue=false;
		//var a=<%=Session("mENURL")%>;
		//}
		//alert("確定離開.")
		//用來引發一個Server端事件
		//document.all("dummy_btn").click();
	}
	//add by ted 20070110
	for (var j=0;j<document.all.length;j++)
	{

		if (document.all(j).tagName=="A")
		{
			document.all(j).disabled = true;
			document.all(j).removeAttribute("href");
		}
		if (document.all(j).tagName=="INPUT")
		{
			/*
			if (document.all(j).type!="hidden")
			{
				if (document.all(j).disabled==false)
				{
					document.all(j).disabled=true;
				}
			}
			*/
		    //還是要判斷一下 因為像type=text如果disabled 則換頁存檔時ViewState會不見
		    if (document.all(j).type == "submit" || document.all(j).type == "button" || document.all(j).type == "reset")
			{
				if (document.all(j).disabled==false)
				{
					document.all(j).disabled=true;
				}
            }

            //if (document.all(j).type == "radio") {
            //    document.all(j).attributes("onclick").value = "return false";
            //}
		}
	}

}

//鎖住非數字鍵
//add by Willy 20081105
//onkeydown="return KCNumCheck();"
function KCNumCheck(){
	if (!window.event.shiftKey) {
		var c = String.fromCharCode(event.keyCode);
		var keyCodeArray = new Array(13,3,8, 9, 35, 36,37,39,45,46,144,145,96,97,98,99,100,101,102,103,104,105);
		for(var i=0;i<=keyCodeArray.length-1;i++){
			if(window.event.keyCode==keyCodeArray[i]){
				return true;
			}
		}
		if ("0123456789".indexOf(c,0)>= 0 ) {
			return true;
		}
	}
	return false;
}

function LockAndMax(text,max){
/*鎖住非數字鍵並限制輸入字數(max)
  用法onkeydown="return LockAndMax(this,5);"
  不限制數字onkeydown="return LockAndMax(this,-1);"
*/
	if (!window.event.shiftKey) {
		var c = String.fromCharCode(event.keyCode);
		if (window.event.keyCode		== KC_BACKSPACE )       {return true;}
		if (window.event.keyCode		== KC_TAB       )       {return true;}
		if (window.event.keyCode		== KC_END       )       {return true;}
		if (window.event.keyCode		== KC_HOME      )       {return true;}
		if (window.event.keyCode		== KC_LEFT      )       {return true;}
		if (window.event.keyCode		== KC_RIGHT     )       {return true;}
		if (window.event.keyCode		== KC_INSERT    )       {return true;}
		if (window.event.keyCode		== KC_DELETE    )       {return true;}
		if (window.event.keyCode		== KC_NUMLOCK   )       {return true;}
		if (window.event.keyCode		== KC_CAPSLOCK  )       {return true;}
		if (window.event.keyCode		== KC_PAD_0     )       {return true;}
		if (window.event.keyCode		== KC_PAD_1     )       {return true;}
		if (window.event.keyCode		== KC_PAD_2     )       {return true;}
		if (window.event.keyCode		== KC_PAD_3     )       {return true;}
		if (window.event.keyCode		== KC_PAD_4     )       {return true;}
		if (window.event.keyCode		== KC_PAD_5     )       {return true;}
		if (window.event.keyCode		== KC_PAD_6     )       {return true;}
		if (window.event.keyCode		== KC_PAD_7     )       {return true;}
		if (window.event.keyCode		== KC_PAD_8     )       {return true;}
		if (window.event.keyCode		== KC_PAD_9     )       {return true;}
		if ("0123456789".indexOf(c,0) 	>= 0			){
			var addascii=0 ;
			var sTemp = text.value + c;
			if (max > 0){
				for(var i=0;i<sTemp.length  ;i++){
					var codes=sTemp.charCodeAt(i);
					if(codes>=32 && codes<=10000){
						addascii++;
					}else{
						addascii+=2;
					}
					if(addascii > max){
						alert( "已輸入超過"+max+"字元長度");
						return false;
					}
				}	
			}					
			return true;
		}
	}
	return false;
}		
// </SCRIPT>
// -->

