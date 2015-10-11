<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="MBSignIn_01_v01.aspx.vb" Inherits="MBSC.MBSignIn_01_v01" %>

<!DOCTYPE html>

<html>
<head>
    <title>會員登入</title>
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
            $('#txt_UserId').focus(function () {
                if ($("#txt_UserId").val() == "帳號(e-Mail)") {
                    $("#txt_UserId").css("color", "black");
                    $("#txt_UserId").val("");
                }
            });
            $('#txt_UserId').blur(function () {
                if ($('#txt_UserId').val() == '') {
                    $("#txt_UserId").css("color", "Gray");
                    $("#txt_UserId").val("帳號(e-Mail)");
                }
            })

            $('#password-clear').show();
            $('#txt_Password').hide();

            $('#password-clear').focus(function () {
                $('#password-clear').hide();
                $('#txt_Password').show();
                $('#txt_Password').focus();
            });
            $('#txt_Password').blur(function () {
                if ($('#txt_Password').val() == '') {
                    $('#password-clear').show();
                    $('#txt_Password').hide();
                }
            });

            if ($("#txt_UserId").val() == "" || $("#txt_UserId").val() == "帳號(e-Mail)") {
                /*每次Dom載入完，確保圖片都不一樣*/
                if (document.all("IMG_Vad")!=undefined)
                {
                    jQuery("img[name='IMG_Vad']").attr("src", "../Module/ValidateNumber.ashx?" + Math.random());
                }                
            }
        });

        function isPassValidateCode() {
            if ($.trim($("#TXT_NUMBER").val()) == "") {
                alert("請輸入驗證碼");
                return false;
            }

            $.ajax({
                type: "POST",
                url: 'MBSignIn_01_v01.aspx/ValidINPUT',
                data: JSON.stringify({
                    'sNUMBER': $("#TXT_NUMBER").val()
                }),
                contentType: "application/json; charset=utf-8",
                async: false,
                dataType: "json",
                success: function (msg) {
                    $("#span_NUMBER").html("");
                    var obj = JSON.parse(msg.d);
                    if (obj.NUMBER == "") {
                        //alert("true");
                        return true;
                    } else {
                        if (obj.NUMBER != "") {
                            $("#span_NUMBER").html(obj.NUMBER);
                            alert(obj.NUMBER);
                        }
                    }
                },
                error: function (jqXHR) {
                    var response = JSON.parse(jqXHR.responseText);
                    alert(response.Message);
                }
            });

            //回傳true Or false
            //alert("false");
            return false;
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
        <!-- #include virtual="~/inc/PageTab.inc" -->
        <!--錯誤訊息區-->
        <!-- #include virtual="~/inc/MBSCErrorMsg.inc" -->
        <!-- #include virtual="~/inc/Signin.inc" -->
        <div class="container">
            <div class="row">
                <div class="col-md-4">
                </div><!-- /.col-md-4-1 -->

                <div class="col-md-4">
                    <div class="panel panel-default">
                        <div class="panel-heading">
                        <div class="panel-title">會員登入</div>
                    </div>
                            
                    <div class="panel-body">
                        <div class="input-group" style="min-width: 100%;">
                                  
					                <asp:TextBox ID="txt_UserId" runat="server" Width="175px" CssClass="mbsctxt" Text="帳號(e-Mail)" Style="color: Gray;height:30px;width:200px" />
					                <br />
					                <input id="password-clear" type="text" value="密碼" autocomplete="off" size="20" class="mbsctxt" style="color: Gray;height:30px;width:200px" />
					                <asp:TextBox ID="txt_Password" runat="server" CssClass="mbsctxt" TextMode="Password" Columns="20" style="height:30px;width:200px" />
					                <br />
					                <br />
					                <asp:Button ID="btLogin" runat="server" Text="登入" CssClass="mbscbt" Style="height:30px;width:200px;font-size:12pt" />
					                <span></span>
					                <br />
					                <br />
					                <asp:Button ID="btSign" runat="server" Text="成為會員" CssClass="mbscbt" Style="height:30px;width:200px;font-size:12pt" />
					                <span></span>
					                <br />
					                <br />
					                <asp:Button ID="btnReMail" runat="server" Text="重發認證信" CssClass="mbscbt" Style="height:30px;width:200px;font-size:12pt" />
					                <span></span>
					                <br />
					                <br />
					                <asp:Button ID="btnFPass" runat="server" Text="忘記密碼" CssClass="mbscbt" Style="height:30px;width:200px;font-size:12pt" />
					                <asp:PlaceHolder ID="PLH_NUMBER" runat="server" Visible="false" >
					                <div style="clear: both;"></div>
					                <div class="rfm">
					                    <table width="100%" cellpadding="0" cellspacing="0" border="0">
					                        <tr>
					                            <td width="100%" valign="top" style="background:transparent" >
					                                <span style="color:red">*</span>
					                                <span style="font-size:11pt;">驗證碼</span>                                         
					                                <asp:TextBox ID="TXT_NUMBER" runat="server" CssClass="mbsctxt" Columns="30" />
					                            </td>
					                        </tr>
					                        <tr>
					                            <td width="100%" valign="top" style="background:transparent;" >
					                                <span style="font-size:11pt;">輸入下圖中的數字</span>
					                                <span style="font-size:9pt;color:red">【看不清楚可點擊下圖更新】</span>
					                                <BR />
					                                <img id="IMG_Vad" name="IMG_Vad" runat="server" src="../Module/ValidateNumber.ashx" alt="驗證碼" />
					                                <span id="span_NUMBER" style="font-size:14pt;font-weight:bold;color:red;" />
					                            </td>
					                        </tr>
					                    </table>
					                    <div class="rfm">
					                        <asp:Button ID="btnReSetPass" runat="server" Text="重設密碼" CssClass="mbscbt" Style="height:30px;width:200px;font-size:12pt" onmousedown="return isPassValidateCode();" />
					                        <div style="font-size:12pt;color:red;font-weight:bold;text-align:left">【輸入正確的驗證碼後，請按重設密碼，系統會寄信至您的Email，請開啟主旨為『MBSC會員重設密碼』的信件，然後點擊其中的重設密碼連結，完成重設密碼程序】</div>
					                    </div>
					                </div>
					                </asp:PlaceHolder>

                            </div><!-- /.btn group -->
                    </div><!-- /.panel-body -->

                     </div><!-- /.panel panel-default-->
                 </div><!-- /.col-md-4-2-->
                 <div class="col-md-4">
                 </div><!-- /.col-md-4-3 -->
             </div><!-- /.row-->
         </div><!-- /.container-->
    </form>
</body>
</html>
