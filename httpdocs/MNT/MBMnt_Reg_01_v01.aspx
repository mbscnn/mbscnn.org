<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="MBMnt_Reg_01_v01.aspx.vb" Inherits="MBSC.MBMnt_Reg_01_v01" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
    <title>註冊MBSC佛陀原始正法中心</title>
    <!-- #include virtual="~/inc/MBSCCSS.inc" -->
    <!-- #include virtual="~/inc/MBSCJS.inc" -->
    <style type="text/css" >
        .rfm
        {
            margin: 0 auto;
            width: 760px;
            border-bottom: 1px dotted black;
            padding: 10px 0px 10px 0px;
        }
        #span_result
        {
            color: Red;
            font-size: 12px;
        }
    </style>
    <script type="text/javascript">
        $(document).ready(function ($) {
            if ($("#HID_USERID").val() == "") {
                /*每次Dom載入完，確保圖片都不一樣*/
                jQuery("img[name='IMG_Vad']").attr("src", "../Module/ValidateNumber.ashx?" + Math.random());
            }
        });

        function isPassValidateCode() {
            if ($.trim($("#TXT_APPNAME").val()) == "") {
                alert("請輸入姓名");
                return false;
            }

            if ($.trim($("#TXT_EMAIL").val()) == "") {
                alert("請輸入EMAIL");
                return false;
            }

            if ($.trim($("#TXT_PASSWORD").val()) == "") {
                alert("請輸入密碼");
                return false;
            }

            if ($.trim($("#TXT_VARIFY").val()) == "") {
                alert("請輸入確認密碼");
                return false;
            }

            if ($.trim($("#TXT_NUMBER").val()) == "") {
                alert("請輸入驗證碼");
                return false;
            }

            if ($.trim($("#TXT_PASSWORD").val()) != $.trim($("#TXT_VARIFY").val())) {
                alert("密碼與確認密碼不符合");
                return false;
            }

            //var isChecked0 = $('#RBL_MB_SEX_0').attr('checked') ? true : false;
            //var isChecked1 = $('#RBL_MB_SEX_1').attr('checked') ? true : false;

            var isChecked0 = false;
            if (document.all("RBL_MB_SEX_0") != undefined)
            {
                if (document.all("RBL_MB_SEX_0").checked) {
                    isChecked0 = true
                }
            }
            var isChecked1 = false;
            if (document.all("RBL_MB_SEX_1") != undefined) {
                if (document.all("RBL_MB_SEX_1").checked) {
                    isChecked1 = true
                }
            }
            if (isChecked0 == false && isChecked1 ==false) {
                alert("請選擇性別");
                return false;
            }

            if ($.trim($("#MB_MOBIL").val()) == "" && $.trim($("#MB_TEL").val()) == "") {
                alert("手機或電話請則一輸入");
                return false;
            }

            $.ajax({
                type: "POST",
                url: 'MBMnt_Reg_01_v01.aspx/ValidINPUT',
                data: JSON.stringify({
                    'sNUMBER':$("#TXT_NUMBER").val() ,
                    'sEMAIL':$("#TXT_EMAIL").val(),
                    'sPASSWORD':$("#TXT_PASSWORD").val(),
                    'sVARPASSWORD': $("#TXT_VARIFY").val(),
                    'sMB_MOBIL': $("#MB_MOBIL").val(),
                    'sMB_TEL_Pre': $("#MB_TEL_Pre").val(),
                    'sMB_TEL': $("#MB_TEL").val()
                }),
                contentType: "application/json; charset=utf-8",
                async: false,
                dataType: "json",
                success: function (msg) {
                    $("#span_NUMBER").html("");
                    $("#span_EMAIL").html("");
                    $("#span_PASSWORD").html("");
                    var obj = JSON.parse(msg.d);
                    if (obj.NUMBER == "" && obj.EMAIL == "" && obj.PASSWORD == "") {
                        return true;
                    } else {
                        var sERR = "";

                        if (obj.NUMBER != "") {
                            $("#span_NUMBER").html(obj.NUMBER);
                            sERR = sERR + obj.NUMBER + "\n";
                        }

                        if (obj.EMAIL != "") {
                            $("#span_EMAIL").html(obj.EMAIL);
                            sERR = sERR + obj.EMAIL + "\n";
                        }

                        if (obj.PASSWORD != "") {
                            $("#span_PASSWORD").html(obj.PASSWORD);
                            sERR = sERR + obj.PASSWORD + "\n";
                        }

                        alert(sERR);
                    }
                },
                error: function (jqXHR) {
                    var response = JSON.parse(jqXHR.responseText);
                    alert(response.Message);
                }
            });

            //回傳true Or false
            return false;
        }

        function isRE_PassValidateCode() {
            if ($.trim($("#TXT_APPNAME").val()) == "") {
                alert("請輸入姓名");
                return false;
            }

            if ($.trim($("#TXT_EMAIL").val()) == "") {
                alert("請輸入EMAIL");
                return false;
            }

            if ($.trim($("#TXT_PASSWORD").val()) == "") {
                alert("請輸入密碼");
                return false;
            }

            if ($.trim($("#TXT_VARIFY").val()) == "") {
                alert("請輸入確認密碼");
                return false;
            }

            if ($.trim($("#TXT_NUMBER").val()) == "") {
                alert("請輸入驗證碼");
                return false;
            }

            if ($.trim($("#TXT_PASSWORD").val()) != $.trim($("#TXT_VARIFY").val())) {
                alert("密碼與確認密碼不符合");
                return false;
            }

            var isChecked0 = $('#RBL_MB_SEX_0').attr('checked') ? true : false;
            var isChecked1 = $('#RBL_MB_SEX_1').attr('checked') ? true : false;
            if (isChecked0 == false && isChecked1 == false) {
                alert("請選擇性別");
                return false;
            }

            if ($.trim($("#MB_MOBIL").val()) == "" && $.trim($("#MB_TEL").val()) == "") {
                alert("手機或電話請則一輸入");
                return false;
            }

            $.ajax({
                type: "POST",
                url: 'MBMnt_Reg_01_v01.aspx/RE_ValidINPUT',
                data: JSON.stringify({
                    'sNUMBER': $("#TXT_NUMBER").val(),
                    'sEMAIL': $("#TXT_EMAIL").val(),
                    'sPASSWORD': $("#TXT_PASSWORD").val(),
                    'sVARPASSWORD': $("#TXT_VARIFY").val()
                }),
                contentType: "application/json; charset=utf-8",
                async: false,
                dataType: "json",
                success: function (msg) {
                    $("#span_NUMBER").html("");
                    $("#span_EMAIL").html("");
                    $("#span_PASSWORD").html("");
                    var obj = JSON.parse(msg.d);
                    if (obj.NUMBER == "" && obj.EMAIL == "" && obj.PASSWORD == "") {
                        return true;
                    } else {
                        var sERR = "";

                        if (obj.NUMBER != "") {
                            $("#span_NUMBER").html(obj.NUMBER);
                            sERR = sERR + obj.NUMBER + "\n";
                        }

                        if (obj.EMAIL != "") {
                            $("#span_EMAIL").html(obj.EMAIL);
                            sERR = sERR + obj.EMAIL + "\n";
                        }

                        if (obj.PASSWORD != "") {
                            $("#span_PASSWORD").html(obj.PASSWORD);
                            sERR = sERR + obj.PASSWORD + "\n";
                        }
                        alert(sERR);
                    }
                },
                error: function (jqXHR) {
                    var response = JSON.parse(jqXHR.responseText);
                    alert(response.Message);
                }
            });

            //回傳true Or false
            return false;
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
        <!-- #include virtual="~/inc/PageTab.inc" -->
        <!-- #include virtual="~/inc/MBSCErrorMsg.inc" -->
        <table width="100%" cellspacing="0" cellpadding="0" style="width:1235px;background:transparent;margin-left: auto; margin-right: auto;" align="center">
            <tr>
                <td style="vertical-align:top;padding:0;text-align:left;background:transparent;width:1050px" >
                    <!-- #include virtual="~/inc/Signin.inc" -->            
                    <input type="hidden" id="HID_USERID" runat="server" />
                    <asp:PlaceHolder ID="PLH_MB_ACCT" runat="server">
                    <BR />
                    <table class="CRTable_Top" width="70%" cellspacing="0" cellpadding="0" align=center>
                        <tr>
                            <td width="100%" class="th1_b">
                                <asp:Literal ID="LTL_TITLE" runat="server" Text="註冊" />                                
                            </td>
                        </tr>
                        <tr>
                            <td width="100%" class="td2_b">
                                <div class="rfm">
                                    <table width="100%" cellpadding="0" cellspacing="0" border="0">
                                        <tr>
                                            <TD width="15%" valign="top" style="background:transparent" >
                                                <span style="color:red">*</span>
                                                <span style="font-size:11pt;">Email</span>
                                            </TD>
                                            <td width="85%" valign="top" style="background:transparent" >
                                                <asp:TextBox ID="TXT_EMAIL" runat="server" CssClass="mbsctxt" Columns="30" />
                                                <span id="span_EMAIL" style="color:red" />
                                            </td>
                                        </tr>
                                    </table>
                                </div>
                                <div class="rfm">
                                    <table width="100%" cellpadding="0" cellspacing="0" border="0">
                                        <tr>
                                            <TD width="15%" valign="top" style="background:transparent" >
                                                <span style="color:red">*</span>
                                                <span style="font-size:11pt;">密碼</span>
                                            </TD>
                                            <td width="85%" valign="top" style="background:transparent" >
                                                <asp:TextBox ID="TXT_PASSWORD" runat="server" CssClass="mbsctxt" Style="IME-MODE: disabled;" Columns="30" TextMode=Password />
                                                <span id="span_PASSWORD" style="color:red" />
                                            </td>
                                        </tr>
                                        <tr>
                                            <td colspan="2" style="color:blue;background:transparent">
                                                *注意密碼最多15碼最少8碼
                                            </td>
                                        </tr>
                                    </table>
                                </div>
                                <div class="rfm">
                                    <table width="100%" cellpadding="0" cellspacing="0" border="0">
                                        <tr>
                                            <TD width="15%" valign="top" style="background:transparent" >
                                                <span style="color:red">*</span>
                                                <span style="font-size:11pt;">確認密碼</span>
                                            </TD>
                                            <td width="85%" valign="top" style="background:transparent" >
                                                <asp:TextBox ID="TXT_VARIFY" runat="server" CssClass="mbsctxt" Style="IME-MODE: disabled;" Columns="30" TextMode=Password />
                                            </td>
                                        </tr>
                                    </table>
                                </div>
                                <div class="rfm">
                                    <table width="100%" cellpadding="0" cellspacing="0" border="0" >
                                        <tr>
                                            <TD width="15%" valign="top" style="background:transparent" >
                                                <span style="color:red">*</span>
                                                <span style="font-size:11pt;">姓名</span>                            
                                            </TD>
                                            <td width="85%" valign="top" style="background:transparent" >
                                                <asp:TextBox ID="TXT_APPNAME" runat="server" CssClass="mbsctxt" Columns="30" />
                                            </td>
                                        </tr>
                                    </table>                        
                                </div>
                                <div class="rfm">
                                    <table width="100%" cellpadding="0" cellspacing="0" border="0" >
                                        <tr>
                                            <TD width="15%" valign="top" style="background:transparent" >
                                                <span style="color:red">*</span>
                                                <span style="font-size:11pt;">性別</span>                            
                                            </TD>
                                            <td width="85%" valign="top" style="background:transparent" >
                                                <asp:RadioButtonList ID="RBL_MB_SEX" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" >
                                                    <asp:ListItem Value="M"><span style="font-size:11pt;">男</span></asp:ListItem>
                                                    <asp:ListItem Value="F"><span style="font-size:11pt;">女</span></asp:ListItem>
                                                </asp:RadioButtonList>
                                            </td>
                                        </tr>
                                    </table>                        
                                </div>
                                <div class="rfm">
                                    <table width="100%" cellpadding="0" cellspacing="0" border="0" >
                                        <tr>
                                            <TD width="15%" valign="top" style="background:transparent" >
                                                <span style="color:red">*</span>
                                                <span style="font-size:11pt;">手機或電話</span>  
                                            </TD>
                                            <td width="85%" valign="top" style="background:transparent" >
                                                <span style="font-size:11pt;">手機</span>
                                                <BR/>
                                                <asp:TextBox ID="MB_MOBIL" runat="server" CssClass="bordernum" style="text-align:left" />
                                                <BR/>
                                                <span style="font-size:11pt;">電話</span>
                                                <BR/>
                                                <asp:TextBox ID="MB_TEL_Pre" runat="server" CssClass="bordernum" Columns="2" />
                                                ─
                                                <asp:TextBox ID="MB_TEL" runat="server" CssClass="bordernum" style="text-align:left" />
                                            </td>
                                        </tr>
                                    </table>
                                </div>                                                           
                                <div class="rfm">
                                    <table width="100%" cellpadding="0" cellspacing="0" border="0">
                                        <tr>
                                            <TD width="10%" valign="top" style="background:transparent" >
                                                <span style="color:red">*</span>
                                                <span style="font-size:11pt;">驗證碼</span>
                                            </TD>
                                            <td width="90%" valign="top" style="background:transparent" > 
                                                <asp:TextBox ID="TXT_NUMBER" runat="server" CssClass="mbsctxt" Columns="30" />
                                    
                                                <BR />
                                                <span style="font-size:11pt;">輸入下圖中的數字</span>
                                                <span style="font-size:9pt;color:red">【看不清楚可點擊下圖更新】</span>
                                                <BR />
                                                <img id="IMG_Vad" name="IMG_Vad" runat="server" src="../Module/ValidateNumber.ashx" alt="驗證碼" />
                                                <span id="span_NUMBER" style="color:red" />                               
                                            </td>
                                        </tr>
                                    </table>
                                </div>                    
                            </td>
                        </tr>
                    </table>
                    
                    <div class="rfm" align="center">
                        <asp:PlaceHolder ID="PLH_Send" runat="server" >
                            <asp:Button ID="btSend" runat="server" CssClass="bt" Text="確定送出" onmousedown="return isPassValidateCode();" />
                            <div >
                                <BR />
                                <ul style="font-size:12pt;color:red;text-align:left" >
                                    <li>確定送出後，請至您的Email，開啟MBSC會員認證信，並點擊其中的認證連結，完成註冊程序</li>
                                    <li>請填寫正確e-mail 及 電話以免影響您未來的活動報名之權益</li>
                                </ul>                                
                            </div>
                        </asp:PlaceHolder>
                        <asp:PlaceHolder ID="PLH_REPASS" runat="server" Visible="false" >
                            <asp:Button ID="bntRePASS" runat="server" CssClass="bt" Text="重設密碼" onmousedown="return isRE_PassValidateCode();" />
                        </asp:PlaceHolder>
                    </div>
                    <br />
                    </asp:PlaceHolder>
                    <asp:PlaceHolder ID="PLH_MAILSend" runat="server" Visible="false">
                        <div align="center" style="font-size:12pt;color:red;font-weight:bold" >
                            <BR />
                            系統已發送驗證至您的e-Mail，請到您的信箱完成帳號啟用<BR/>
                            【若找不到驗證信，請試試垃圾郵件資料匣】
                        </div>
                    </asp:PlaceHolder>
                    <asp:PlaceHolder ID="PLH_APV" runat="server" Visible="false">
                        <div align="center" style="font-size:12pt;color:red;font-weight:bold" >
                            <BR />
                            您的帳號已經啟用
                        </div>
                        <asp:Literal ID="LTL_SCRIPT" runat="server" Visible="false" />
                    </asp:PlaceHolder>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
