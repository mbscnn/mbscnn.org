<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="MBMnt_Member_01_v01.aspx.vb"
    Inherits="MBSC.MBMnt_Member_01_v01" EnableEventValidation="false" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>法友會會員—入會申請單</title>
    <!-- #include virtual="~/inc/MBSCCSS.inc" -->
    <!--CSS-->
    <!-- #include virtual="~/inc/MBSCJS.inc" -->
    <!--JS-->
    <script language="javascript">
        function YYYCheck() {
            if (window.event.srcElement.value.length == 0) {
                return true;
            }
            if (isNaN(window.event.srcElement.value)) {
                window.event.srcElement.select();
                window.event.srcElement.focus();
                alert("請輸入數字!");
                return false;
            }
            var D_Date = new Date();

            if (KCParseDec(window.event.srcElement.value) > D_Date.getFullYear()) {
                window.event.srcElement.select();
                window.event.srcElement.focus();
                alert("出生年不可大於今年!");
                return false;
            }

            return true;
        }

        $(document).ready(function () {
            $("#MB_CITY").change(function () {
                if ($("#MB_CITY>option:selected").get(0).text.length == 0) { return; }

                var data = { 'sCity': $("#MB_CITY>option:selected").get(0).text };

                $.ajax({
                    type: "POST",
                    url: 'MBMnt_Member_01_v01.aspx/Bind_DDL_Town',
                    data: JSON.stringify(data),
                    contentType: "application/json; charset=utf-8",
                    async: false,
                    dataType: "json",
                    success: function (msg) {
                        $("#MB_VLG").empty();
                        $("<option value='" + "" + "'>" + "《鄉鎮市區》" + "</option>").appendTo("#MB_VLG");
                        var obj = JSON.parse(msg.d);
                        $.each(obj.AddrCode, function (n, value) {
                            $("<option value='" + value.AREA_ID + "'>" + value.AREA + "</option>").appendTo("#MB_VLG");
                        });

                        if (document.all("MB_ADDR2_SAME").change == true)
                        {
                            //alert($('option:selected', '#MB_CITY').index());
                            //$('#MB_CITY1 option')[$('option:selected', '#MB_CITY').index()].selected = true;

                            //$("#MB_VLG1").empty();
                            //$("<option value='" + "" + "'>" + "《鄉鎮市區》" + "</option>").appendTo("#MB_VLG1");
                            //$.each(obj.AddrCode, function (n, value) {
                            //    $("<option value='" + value.AREA_ID + "'>" + value.AREA + "</option>").appendTo("#MB_VLG1");
                            //});
                            doSame();
                        }
                    },
                    error: function (jqXHR) {
                        var response = JSON.parse(jqXHR.responseText);
                        alert(response.Message);
                    }
                });

                var data2 = { 'sCITY_ID': $("#MB_CITY>option:selected").get(0).value };
                $("#MB_LEADER").empty();
                $('#MB_AREA option')[0].selected = true;
                $.ajax({

                    type: "POST",
                    url: 'MBMnt_Member_01_v01.aspx/Bind_MB_AREA',
                    data: JSON.stringify(data2),
                    contentType: "application/json; charset=utf-8",
                    async: false,
                    dataType: "json",
                    success: function (msg) {
                        var obj = JSON.parse(msg.d);

                        $("#MB_AREA option[value=" + obj.GA_AREA.GA_AREA + "]").attr("selected", "selected");

                        $("<option value='" + "" + "'>" + "請選擇" + "</option>").appendTo("#MB_LEADER");
                        $.each(obj.MB_LEADER, function (n, value) {
                            $("<option value='" + value.MB_LEADER_VALUE + "'>" + value.MB_LEADER_TEXT + "</option>").appendTo("#MB_LEADER");
                        });

                        setTimeout(function () { $("#MB_LEADER option[value=" + $("#HID_MB_LEADER").get(0).value + "]").attr("selected", "selected"); }, 1);
                    },
                    error: function (jqXHR) {
                        var response = JSON.parse(jqXHR.responseText);
                        alert(response.Message);
                    }
                });
            });

            $("#MB_CITY1").change(function () {
                if ($("#MB_CITY1>option:selected").get(0).text.length == 0) { return; }
                var data = {
                    'sCity': $("#MB_CITY1>option:selected").get(0).text
                };
                $.ajax({
                    type: "POST",
                    url: 'MBMnt_Member_01_v01.aspx/Bind_DDL_Town',
                    data: JSON.stringify(data),
                    contentType: "application/json; charset=utf-8",
                    async: false,
                    dataType: "json",
                    success: function (msg) {
                        $("#MB_VLG1").empty();
                        $("<option value='" + "" + "'>" + "《鄉鎮市區》" + "</option>").appendTo("#MB_VLG1");
                        var obj = JSON.parse(msg.d);
                        $.each(obj.AddrCode, function (n, value) {
                            $("<option value='" + value.AREA_ID + "'>" + value.AREA + "</option>").appendTo("#MB_VLG1");
                        });

                        if (document.all("MB_ADDR2_SAME").checked == true)
                        {
                            //alert($('option:selected', '#MB_CITY').index());
                            //$('#MB_CITY1 option')[$('option:selected', '#MB_CITY').index()].selected = true;

                            //$("#MB_VLG1").empty();
                            //$("<option value='" + "" + "'>" + "《鄉鎮市區》" + "</option>").appendTo("#MB_VLG");
                            //$.each(obj.AddrCode, function (n, value) {
                            //    $("<option value='" + value.AREA_ID + "'>" + value.AREA + "</option>").appendTo("#MB_VLG1");
                            //});
                            doSame();
                        }
                    },
                    error: function (jqXHR) {
                        var response = JSON.parse(jqXHR.responseText);
                        alert(response.Message);
                    }
                });

                if (document.all("MB_ADDR2_SAME").checked==true) {
                    var data2 = { 'sCITY_ID': $("#MB_CITY1>option:selected").get(0).value };
                    $("#MB_LEADER").empty();
                    $('#MB_AREA option')[0].selected = true;
                    $.ajax({

                        type: "POST",
                        url: 'MBMnt_Member_01_v01.aspx/Bind_MB_AREA',
                        data: JSON.stringify(data2),
                        contentType: "application/json; charset=utf-8",
                        async: false,
                        dataType: "json",
                        success: function (msg) {
                            var obj = JSON.parse(msg.d);

                            $("#MB_AREA option[value=" + obj.GA_AREA.GA_AREA + "]").attr("selected", "selected");

                            $("<option value='" + "" + "'>" + "請選擇" + "</option>").appendTo("#MB_LEADER");
                            $.each(obj.MB_LEADER, function (n, value) {
                                $("<option value='" + value.MB_LEADER_VALUE + "'>" + value.MB_LEADER_TEXT + "</option>").appendTo("#MB_LEADER");
                            });

                            setTimeout(function () { $("#MB_LEADER option[value=" + $("#HID_MB_LEADER").get(0).value + "]").attr("selected", "selected"); }, 1);
                        },
                        error: function (jqXHR) {
                            var response = JSON.parse(jqXHR.responseText);
                            alert(response.Message);
                        }
                    });
                }
            });

            $("#MB_ID").blur(function () {
                if ($("#MB_ID").get(0).value.length == 0) { return true; }
                var data = {
                    'sMB_ID': $("#MB_ID").get(0).value
                };
                $.ajax({
                    type: "POST",
                    url: 'MBMnt_Member_01_v01.aspx/CHK_MB_ID',
                    data: JSON.stringify(data),
                    contentType: "application/json; charset=utf-8",
                    async: false,
                    dataType: "json",
                    success: function (msg) {
                        var obj = JSON.parse(msg.d);

                        if (obj.CHK_MB_ID != "TRUE") {
                            //$("#MB_ID").get(0).select();
                            if ($('#MB_ID').length > 0) { setTimeout(function () { try { $('#MB_ID')[0].select(); $('#MB_ID')[0].focus(); } catch (e) { } }, 0); }
                            //window.event.srcElement.select();
                            //window.event.srcElement.focus();
                            alert(obj.CHK_MB_ID);
                            return false;
                        }
                        else {
                            $('input[name="MB_SEX"]').removeAttr('checked');
                            $("[name=MB_SEX]").filter("[value=" + obj.MB_SEX + "]").prop("checked", true);

                            return true;
                        }
                    },
                    error: function (jqXHR) {
                        var response = JSON.parse(jqXHR.responseText);
                        alert(response.Message);
                    }
                });
            });
            /*$("#RBL_MEMREV_MB_FAMILY").change(function () {
                //alert($("#RBL_MEMREV_MB_FAMILY input:checked").val());
                switch ($("#RBL_MEMREV_MB_FAMILY input:checked").val()) {
                    case "1":
                        for (var i=0;i<$("#MB_FEETYPE input").length;i++)
                        {
                            var sFEETYPE = "";
                            sFEETYPE = $("#MB_FEETYPE input")[i].value;
                            if (sFEETYPE == "A" || sFEETYPE == "B" || sFEETYPE == "C" || sFEETYPE=="D")
                            {
                                $("#MB_FEETYPE input")[i].style.display = "";
                                $("#MB_FEETYPE input")[i].nextSibling.style.display = "";
                            }
                            else
                            {   
                                $("#MB_FEETYPE input")[i].style.display = "none";
                                $("#MB_FEETYPE input")[i].nextSibling.style.display = "none";
                            }
                        }
                        break;
                    case "2":
                        for (var i = 0; i < $("#MB_FEETYPE input").length; i++) {
                            var sFEETYPE = "";
                            sFEETYPE = $("#MB_FEETYPE input")[i].value;
                            if (sFEETYPE == "E" || sFEETYPE == "F")
                            {
                                $("#MB_FEETYPE input")[i].style.display = "";
                                $("#MB_FEETYPE input")[i].nextSibling.style.display = "";
                            } else {
                                $("#MB_FEETYPE input")[i].style.display = "none";
                                $("#MB_FEETYPE input")[i].nextSibling.style.display = "none";
                            }
                        }
                       break;
                }
            });*/
            $("#MB_FEETYPE").change(function () {
                //alert($("#MB_ITEMID").val());
                var sMB_ITEMID = $("#MB_ITEMID").val();
                if (sMB_ITEMID.length == 0) {
                    alert("請先選擇功德項目");

                    return false;
                }

                //if ($("#MB_FEETYPE input:checked").val() == "A" || $("#MB_FEETYPE input:checked").val() == "B" || $("#MB_FEETYPE input:checked").val() == "C")
                //{
                //    if ($.isNumeric($("#MB_MEMFEE_SY").val()) == false || $.isNumeric($("#MB_MEMFEE_SM").val()) == false)
                //    {
                //        alert("請先輸入會費期間起日");

                //        return false;
                //    }
                //}

                $("#MB_MEMFEE_EY").attr("readonly", false);
                $("#MB_MEMFEE_EM").attr("readonly", false);
                if ($("#MB_FEETYPE input:checked").val() == "D" || $("#MB_FEETYPE input:checked").val() == "E") {
                    $("#MB_TOTFEE").attr("readonly", false);
                    $("#MB_TOTFEE").val("");
                    $("#MB_TOTFEE_MM").val("");
                    $("#PLH_MB_TOTFEE_MM").css("display", "none");

                    if ($("#MB_FEETYPE input:checked").val() == "E")
                    {
                        $("#ACCAMT").val("");
                    }
                }
                else if ($("#MB_FEETYPE input:checked").val() == "F") {
                    $("#MB_TOTFEE").attr("readonly", true);
                    $("#MB_TOTFEE").val("100,000");
                    $("#MB_TOTFEE_MM").val("");
                    $("#PLH_MB_TOTFEE_MM").css("display", "none");
                    var iACCAMT = 0
                    if ($.isNumeric($("#HID_ACCAMT").val())) {
                        iACCAMT = $("#HID_ACCAMT").val() * 1
                    }

                    $("#ACCAMT").val(100000 + iACCAMT);
                }
                else
                {
                    if ($("#MB_FEETYPE input:checked").val() == "A")
                    {
                        //$("#MB_MEMFEE_EY").val($("#MB_MEMFEE_SY").val());
                        //$("#MB_MEMFEE_EM").val($("#MB_MEMFEE_SM").val());
                        //$("#MB_MEMFEE_EY").attr("readonly", true);
                        //$("#MB_MEMFEE_EM").attr("readonly", true);
                    }
                    else if ($("#MB_FEETYPE input:checked").val() == "B")
                    {
                        //var sSDate = (parseInt($("#MB_MEMFEE_SY").val())) + "/" + $("#MB_MEMFEE_SM").val() + "/1";
                        //var D = new Date(sSDate);
                        //D.setMonth(D.getMonth() + 3);
                        //$("#MB_MEMFEE_EY").val(D.getFullYear());
                        //$("#MB_MEMFEE_EM").val(D.getMonth()+1);
                        //$("#MB_MEMFEE_EY").attr("readonly", true);
                        //$("#MB_MEMFEE_EM").attr("readonly", true);
                    }
                    else if ($("#MB_FEETYPE input:checked").val() == "C") {
                        //var sSDate = (parseInt($("#MB_MEMFEE_SY").val())) + "/" + $("#MB_MEMFEE_SM").val() + "/1";
                        //var D = new Date(sSDate);
                        //D.setMonth(D.getMonth() + 12);
                        //$("#MB_MEMFEE_EY").val(D.getFullYear());
                        //$("#MB_MEMFEE_EM").val(D.getMonth()+1);
                        //$("#MB_MEMFEE_EY").attr("readonly", true);
                        //$("#MB_MEMFEE_EM").attr("readonly", true);
                    }

                    $("#MB_TOTFEE").val("");
                    var sMB_FEETYPE = { 'sMB_FEETYPE': $("#MB_FEETYPE input:checked").val() };
                    $.ajax({
                        type: "POST",
                        url: 'MBMnt_Member_01_v01.aspx/get23_NOTE',
                        data: JSON.stringify(sMB_FEETYPE),
                        contentType: "application/json; charset=utf-8",
                        async: false,
                        dataType: "json",
                        success: function (msg) {
                            if (msg != undefined) {
                                //alert(msg.d);
                                $("#MB_TOTFEE").val(msg.d);
                            }
                        },
                        error: function (jqXHR) {
                            var response = JSON.parse(jqXHR.responseText);
                            alert(response.Message);
                        }
                    });
                    
                    $("#MB_TOTFEE").attr("readonly", true);
                    $("#MB_TOTFEE_MM").val("");
                    //$("#PLH_MB_TOTFEE_MM").css("display", "");
                    $("#PLH_MB_TOTFEE_MM").css("display", "none");
                }
            });
            $("#MB_TOTFEE_MM").bind("input propertychange", function () {
                //alert($.isNumeric($("#MB_TOTFEE_MM").val().replace(/,/g,"")));
                if ($.isNumeric($("#MB_TOTFEE_MM").val().replace(/,/g,"")))
                {
                    if ($("#MB_FEETYPE input:checked").val() == "B")
                    {
                        $("#MB_TOTFEE").val(KCParseDec($("#MB_TOTFEE_MM").val()) * 3);
                    }
                    else if ($("#MB_FEETYPE input:checked").val() == "C") {
                        $("#MB_TOTFEE").val(KCParseDec($("#MB_TOTFEE_MM").val()) * 12);
                    }
                    else if ($("#MB_FEETYPE input:checked").val() == "A") {
                        $("#MB_TOTFEE").val(KCParseDec($("#MB_TOTFEE_MM").val()) * 1);
                    }
                }
                else
                {
                    if ($("#MB_FEETYPE input:checked").val() != "F") {
                        $("#MB_TOTFEE").val("");
                    }
                }
            });
        });

        function doSame() {
            if (document.all("MB_ADDR2_SAME").checked == true)
            {
                //alert($('option:selected', '#MB_CITY').index());
                $('#MB_CITY1 option')[$('option:selected', '#MB_CITY').index()].selected = true;
                $("#MB_VLG1").empty();

                $("#MB_VLG").children().each(function () {
                    $("<option value='" + $(this).val() + "'>" + $(this).text() + "</option>").appendTo("#MB_VLG1");
                });

                $("#MB_ADDR2").val($("#MB_ADDR").val());
                $("#TXT_MB_ZIP1").val($("#TXT_MB_ZIP").val());

                var iIndex = $('option:selected', '#MB_VLG').index();

                setTimeout(function () { $('#MB_VLG1 option')[iIndex].selected = true; }, 1);
            }
        }

        function MB_VLG_onChange() {
            //if (document.all("MB_ADDR2_SAME").checked==true) {
            //    $('#MB_VLG1 option')[$('option:selected', '#MB_VLG').index()].selected = true;
            //}
            doSame();
        }

        function MB_VLG1_onChange() {
            //if (document.all("MB_ADDR2_SAME").checked == true) {
            //    $('#MB_VLG1 option')[$('option:selected', '#MB_VLG').index()].selected = true;
            //}
            doSame();
        }

        function MB_ADDR_onChange() {
            //if (document.all("MB_ADDR2_SAME").checked == true) {
            //    $("#MB_ADDR2").val($("#MB_ADDR").val());
            //}
            doSame();
        }

        function MB_ADDR2_onChange() {
            //if (document.all("MB_ADDR2_SAME").checked == true) {
            //    $("#MB_ADDR2").val($("#MB_ADDR").val());
            //}
            doSame();
        }

        function MB_ITEMID_onchange(thisObj) {
            //alert(thisObj.value);
            if (thisObj.value.length == 0) {
                alert("請選擇功德項目!");
                return;
            }
            if (thisObj.value == "A") {
                //會員
                $("#PLH_MB_FAMILY").css("display", "");
                $("#TR_FUND").css("display", "");
                $("#MB_FEETYPE").css("display", "");
                $("#LBL_ONLY").css("display", "none");
                var radiolist_MB_FEETYPE = $("#MB_FEETYPE").find('input:radio');
                radiolist_MB_FEETYPE.removeAttr('checked');
                $("#LTL_MB_ITEMID_TYP").html("會費期間");
                $("#PLH_MB_MEMFEE").css("display", "");
                $("#MB_DESC").css("display", "none");
                $('#MB_TOTFEE').attr("readonly", true);
                $("#PLH_MB_TOTFEE_MM").css("display", "none");
                $("#PLH_PROJCODE").css("display", "none");
            }
            else {
                //非會員
                $("#PLH_MB_FAMILY").css("display", "none");;
                $("#TR_FUND").css("display", "none");
                var radiolist = $("#RBL_MEMREV_MB_FAMILY").find('input:radio');
                radiolist.removeAttr('checked');
                $("#MB_FEETYPE").css("display", "none");
                $("#LBL_ONLY").css("display", "");
                $("#MB_FEETYPE").find("input[value='D']").attr("checked", "checked");
                $("#LTL_MB_ITEMID_TYP").html("說明");
                $("#PLH_MB_MEMFEE").css("display", "none");
                $("#MB_MEMFEE_SY").val("");
                $("#MB_MEMFEE_SM").val("");
                $("#MB_MEMFEE_EY").val("");
                $("#MB_MEMFEE_EM").val("");
                $("#MB_DESC").css("display", "");
                $('#MB_TOTFEE').attr("readonly", false);
                $("#PLH_MB_TOTFEE_MM").css("display", "none");
                if (thisObj.value == "B")
                {
                    $("#PLH_PROJCODE").css("display", "");
                }
            }
        }

        function CB_MB_RECNAME_onclick(thisObj) {
            if (thisObj.checked) {
                $("#MB_RECNAME").val($("#LTL_MEMREV_MB_NAME").html());
            }
            else {
                $("#MB_RECNAME").val("");
            }
        }

        function MB_PAY_TYPE_onclick(thisObj) {
            //alert($("#MB_PAY_TYPE input:checked").val());
            if ($("#MB_PAY_TYPE input:checked").val() == "N") {
                $("#PLH_MB_PAY_TYPE_N").css("display", "");
            }
            else {
                $("#PLH_MB_PAY_TYPE_N").css("display", "none");
                $("#NOTE_DUE_DATE_YYY").val("");
                $("#NOTE_DUE_DATE_MM").val("");
                $("#NOTE_DUE_DATE_DD").val("");
                $("#NOTE_NO").val("");
                $("#NOTE_BANK").val("");
                $("#NOTE_BR").val("");
                $("#NOTE_HOLDER").val("");
                $("#NOTE_AMT").val("");
            }
        }

        function MB_TOTFEE_onchange(thisObj)
        {
            //alert($.isNumeric(thisObj.value));
            if ($("#MB_FEETYPE input:checked").val() == "E")
            {
                if ($.isNumeric(thisObj.value))
                {
                    var iACCAMT = 0

                    if ($.isNumeric($("#HID_ACCAMT").val()))
                    {
                        iACCAMT = $("#HID_ACCAMT").val() * 1
                    }

                    $("#ACCAMT").val(thisObj.value * 1 + iACCAMT);
                }
            }
        }

        function MB_MEMFEE_SYM_CHANGE()
        {
            $("#MB_MEMFEE_EY").attr("readonly", false);
            $("#MB_MEMFEE_EM").attr("readonly", false);
            if ($.isNumeric($("#MB_MEMFEE_SY").val()) == true && $.isNumeric($("#MB_MEMFEE_SM").val()) == true)
            {
                if ($("#MB_FEETYPE input:checked").val() == "A")
                {
                    $("#MB_MEMFEE_EY").val($("#MB_MEMFEE_SY").val());
                    $("#MB_MEMFEE_EM").val($("#MB_MEMFEE_SM").val());
                }
                else if ($("#MB_FEETYPE input:checked").val() == "B")
                {
                    var sSDate = (parseInt($("#MB_MEMFEE_SY").val())) + "/" + $("#MB_MEMFEE_SM").val() + "/1";
                    var D = new Date(sSDate);
                    D.setMonth(D.getMonth() + 3);
                    $("#MB_MEMFEE_EY").val(D.getFullYear());
                    $("#MB_MEMFEE_EM").val(D.getMonth()+1);
                    $("#MB_MEMFEE_EY").attr("readonly", true);
                    $("#MB_MEMFEE_EM").attr("readonly", true);
                }
                else if ($("#MB_FEETYPE input:checked").val() == "C") {
                    var sSDate = (parseInt($("#MB_MEMFEE_SY").val())) + "/" + $("#MB_MEMFEE_SM").val() + "/1";
                    var D = new Date(sSDate);
                    D.setMonth(D.getMonth() + 12);
                    $("#MB_MEMFEE_EY").val(D.getFullYear());
                    $("#MB_MEMFEE_EM").val(D.getMonth()+1);
                    $("#MB_MEMFEE_EY").attr("readonly", true);
                    $("#MB_MEMFEE_EM").attr("readonly", true);
                }
            }
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <asp:PlaceHolder ID="PLH_PAGETAB" runat="server">
    <!-- #include virtual="~/inc/PageTab.inc" -->
    </asp:PlaceHolder>
<table width="100%" cellspacing="0" cellpadding="0" style="width:1235px;background:transparent;margin-left: auto; margin-right: auto;" align="center">
    <tr>
        <td style="vertical-align:top;padding:0;text-align:left;background:transparent;width:1050px" >
            <!-- #include virtual="~/inc/Signin.inc" -->
            <div align="center">
                <h2 style="color: #800080; font-family: 標楷體">
                    社團法人台灣佛陀原始正法學會
                    <br />
                    <asp:Literal ID="LTL_PROGTITLE" runat="server" />
                </h2>
            </div>
            <!--顯示上方外框線-->
            <!-- #include virtual="~/inc/MBSCTableStart.inc" -->
            <!--錯誤訊息區-->
            <!-- #include virtual="~/inc/MBSCErrorMsg.inc" -->
            <asp:PlaceHolder ID="PLH_QRY" runat="server">        
                <table class="CRTable_Top" width="100%" cellspacing="0">
                    <tr>
                        <td width="100%" class="th1_b">
                            請先輸入
                        </td>
                    </tr>
                    <tr>
                        <td width="100%" class="th1_b">
                            <asp:RadioButton ID="RB_QRY_1" runat="server" GroupName="RB_QRY" />
                            手機:
                            <asp:TextBox ID="txt_QRY_Mobile" runat="server" CssClass="bordernum" style="text-align:left" />
                            <asp:RadioButton ID="RB_QRY_2" runat="server" GroupName="RB_QRY" />
                            電話:
                            <asp:TextBox ID="txt_QRY_Phone_Pre" runat="server" CssClass="bordernum" Columns="2" style="text-align:left" />
                            ─
                            <asp:TextBox ID="txt_QRY_Phone" runat="server" CssClass="bordernum" style="text-align:left" />
                            <asp:RadioButton ID="RB_QRY_3" runat="server" GroupName="RB_QRY" />
                            姓名:
                            <asp:TextBox ID="txt_QRY_MB_NAME" runat="server" CssClass="bordertxt" style="text-align:left" />
                        </td>
                    </tr>
                </table>
                <div align="center">
                    <asp:Button ID="bt_Qry" runat="server" CssClass="bt" Text="確　　定" />
                </div>
            </asp:PlaceHolder>

            <asp:PlaceHolder ID="PLH_QRY_SAME" runat="server" Visible=false> 
                <table class="CRTable_Top" width="100%" cellspacing="0">
                    <tr>
                        <td width="100%" class="th1_b" colspan=3>
                            <span style="color:red;font-weight:bold">請點選姓名連結會員資料</span>
                        </td> 
                    </tr>
                    <tr>
                        <td width="25%" class="th1c_b">
                            姓名
                        </td> 
                        <td width="65%" class="th1c_b">
                            通訊地址
                        </td>                
                        <td width="10%" class="th1c_b">
                            會員編號
                        </td>                                               
                    </tr>
                    <asp:Repeater ID="RP_QRY_SAME" runat="server">
                        <ItemTemplate>
                            <tr>
                                <td width="25%" class="td2c_b">
                                    <!--姓名-->
                                    <asp:LinkButton ID="MB_NAME" runat="server" style="text-decoration:underline;font-weight:bold" CommandArgument='<%#Container.DataItem("MB_MEMSEQ")%>' Text='<%#Container.DataItem("MB_NAME")%>'></asp:LinkButton>
                                </td> 
                                <td width="65%" class="td2_b">
                                    <!--通訊地址-->
                                    <%# Container.DataItem("MB_CITY")%><%# Container.DataItem("MB_VLG")%><%#Container.DataItem("MB_ADDR")%>
                                </td>                
                                <td width="10%" class="td2c_b">
                                    <!--會員編號-->
                                    <%# getMB_MEMSEQ(Container.DataItem("MB_MEMSEQ"), Container.DataItem("MB_AREA"))%>
                                </td>                                               
                            </tr>                    
                        </ItemTemplate>
                    </asp:Repeater>
                </table>
                <div align="center">
                    <asp:Button ID="bt_Qry_NO" runat="server" Text="以上皆非" CssClass="bt" />
                </div>
            </asp:PlaceHolder>

            <asp:PlaceHolder ID="PLH_MAIL_SAME" runat="server" Visible="false">
                <table class="CRTable_Top" width="100%" Cellspacing="0">
                    <tr>
                        <td width="15%" class="th1c_b">
                            點選
                        </td>
                        <td width="20%" class="th1c_b">
                            姓名
                        </td>
                        <td width="35%" class="th1c_b">
                            通訊地址
                        </td>
                        <td width="20%" class="th1c_b">
                            會員編號
                        </td>
                    </tr>
                    <asp:Repeater ID="RP_MAIL_SAME" runat="server">
                        <ItemTemplate>
                            <tr>
                                <td width="15%" class="td2c_b">
                                    <!--點選-->
                                    <asp:RadioButton ID="RB_CHOOSE" runat="server" AutoPostBack="true" OnCheckedChanged="RB_CHOOSE_OnCheckedChanged" />
                                    <input type="hidden" id="HID_MB_MEMSEQ" runat="server" value='<%#Container.DataItem("MB_MEMSEQ")%>' />
                                </td>
                                <td width="20%" class="td2c_b">
                                    <!--姓名-->
                                    <%#Container.DataItem("MB_NAME")%>
                                </td>
                                <td width="35%" class="td2_b">
                                    <!--通訊地址-->
                                    <%# Container.DataItem("MB_CITY")%><%# Container.DataItem("MB_VLG")%><%#Container.DataItem("MB_ADDR")%>
                                </td>
                                <td width="20%" class="td2c_b">
                                    <!--會員編號-->
                                    <%# getMB_MEMSEQ(Container.DataItem("MB_MEMSEQ"), Container.DataItem("MB_AREA"))%>
                                </td>
                            </tr>
                        </ItemTemplate>
                    </asp:Repeater>
                </table>
                <div align="center" >
                    <asp:Button ID="btnOTHSign" runat="server" CssClass="bt" Text="幫他人入會" />
                    &nbsp;
                    <asp:Button ID="btnModify" runat="server" CssClass="bt" Text="修改資料" />
                    <input type="hidden" id="HID_OTHSIGN" runat="server" />
                </div>
            </asp:PlaceHolder>

            <asp:PlaceHolder ID="PLH_MB_MEMBER" runat="server" Visible=false >
                <span style="font-weight:bold;font-size:14pt" >
                    會員編號:<asp:Literal ID="LTL_MB_MEMSEQ" runat="server" />
                </span>
                <div style="font-weight:bold;font-size:14pt;color:red">
                    台銀編號:<asp:Literal ID="MB_BKSEQ" runat="server" ></asp:Literal>
                </div>
                <table class="CRTable_Top" width="100%" cellspacing="0">
                    <tr>
                        <td width="15%" class="th1c_b">
                            <span style="color: red; font-weight: bold; font-size: medium">*</span>&nbsp;法名/姓名
                        </td>
                        <td width="35%" class="td2_b">
                            <asp:TextBox ID="MB_NAME" runat="server" CssClass="bordertxt" />
                            <input type="hidden" id="HID_MB_MEMSEQ" runat="server" />
                        </td>
                        <td width="15%" class="th1c_b">
                            <span style="color: red; font-weight: bold; font-size: medium">*</span>&nbsp;性別
                        </td>
                        <td width="35%" class="td2_b">
                            <asp:RadioButtonList ID="MB_SEX" name="MB_SEX" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                <asp:ListItem Text="男" Value="1"></asp:ListItem>
                                <asp:ListItem Text="女" Value="2"></asp:ListItem>
                            </asp:RadioButtonList>
                        </td>
                    </tr>
                    <tr>
                        <td width="15%" class="th1c_b">
                            出生年月日
                        </td>
                        <td width="35%" class="td2_b">
                            西元
                            <asp:TextBox ID="MB_BIRTH_YYY" runat="server" CssClass="bordernum" Columns="4" onfocusout="YYYCheck();" />
                            年
                            <asp:TextBox ID="MB_BIRTH_MM" runat="server" CssClass="bordernum" Columns="2" onfocusout="KCMmCheck();" />
                            月
                            <asp:TextBox ID="MB_BIRTH_DD" runat="server" CssClass="bordernum" Columns="2" onfocusout="KCDdCheck();" />
                            日
                        </td>
                        <td width="15%" class="th1c_b">
                            身分別
                        </td>
                        <td width="35%" class="td2_b">
                            <asp:CheckBoxList ID="MB_IDENTIFY" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                <asp:ListItem Text="會員" Value="A"></asp:ListItem>
                                <asp:ListItem Text="委員" Value="B"></asp:ListItem>
                                <asp:ListItem Text="學員" Value="C"></asp:ListItem>
                                <asp:ListItem Text="法工" Value="D"></asp:ListItem>
                            </asp:CheckBoxList>
                        </td>
                    </tr>
                    <tr>
                        <td width="15%" class="th1c_b">
                            <span style="color: red; font-weight: bold; font-size: medium">*</span>&nbsp;手機
                        </td>
                        <td width="35%" class="td2_b">
                            <asp:TextBox ID="MB_MOBIL" runat="server" CssClass="bordernum" style="text-align:left" />
                            (手機或電話請擇一輸入)
                        </td>
                        <td width="15%" class="th1c_b">
                            <span style="color: red; font-weight: bold; font-size: medium">*</span>&nbsp;電話
                        </td>
                        <td width="35%" class="td2_b">
                            <asp:TextBox ID="MB_TEL_Pre" runat="server" CssClass="bordernum" Columns="2" />
                            ─
                            <asp:TextBox ID="MB_TEL" runat="server" CssClass="bordernum" style="text-align:left" />
                            (手機或電話請擇一輸入)
                        </td>
                    </tr>
                    <tr>
                        <td width="15%" class="th1c_b">
                            e-mail
                        </td>
                        <td width="35%" class="td2_b">
                            <asp:TextBox ID="MB_EMAIL" runat="server" CssClass="bordertxt" onfocusout="if (!isEMailAddr(this)){this.select();this.focus();}" />
                        </td>
                        <td width="15%" class="th1c_b">
                            身分證字號
                        </td>
                        <td width="35%" class="td2_b">
                            <asp:TextBox ID="MB_ID" runat="server" CssClass="bordertxt" />
                        </td>
                    </tr>
                    <tr>
                        <td width="15%" class="th1c_b">
                            學歷
                        </td>
                        <td id="TD_MB_AREA_1" runat="server" width="35%" class="td2_b">
                            <asp:DropDownList ID="MB_EDU" runat="server">
                            </asp:DropDownList>
                        </td>
                        <td id="TD_MB_AREA_2" runat="server" width="15%" class="th1c_b">
                            所屬區
                        </td>
                        <td id="TD_MB_AREA_3" runat="server" width="35%" class="td2_b">
                            <asp:DropDownList ID="MB_AREA" runat="server" onfocus="this.defaultIndex=this.selectedIndex;" onchange="this.selectedIndex=this.defaultIndex;">
                            </asp:DropDownList>
                            &nbsp;委員
                            <asp:DropDownList ID="MB_LEADER" runat="server" onchange='$("#HID_MB_LEADER").val($(this).val());'>
                            </asp:DropDownList>
                            <input type="hidden" id="HID_MB_LEADER" runat="server" />
                        </td>
                    </tr>
                    <tr>
                        <td width="15%" class="th1c_b">
                            <span style="color: red; font-weight: bold; font-size: medium">*</span>&nbsp;通訊地址
                        </td>
                        <td width="85%" class="td2_b" colspan="3">
                            <asp:DropDownList ID="MB_CITY" runat="server">
                            </asp:DropDownList>
                            市
                            <asp:DropDownList ID="MB_VLG" runat="server" onChange="MB_VLG_onChange();"  >
                                <asp:ListItem Text="《鄉鎮市區》" Value=""></asp:ListItem>
                            </asp:DropDownList>
                            區
                            <asp:TextBox ID="MB_ADDR" runat="server" CssClass="bordertxt" width="60%" onChange="MB_ADDR_onChange();" />
                            郵遞區號 :
                            <asp:TextBox ID="TXT_MB_ZIP" runat="server" Columns="3" CssClass="bordernum" />
                        </td>
                    </tr>
                    <tr>
                        <td width="15%" class="th1c_b">
                            <span style="color: red; font-weight: bold; font-size: medium">*</span>&nbsp;戶籍地址
                        </td>
                        <td width="85%" class="td2_b" colspan="3">
                            <asp:CheckBox ID="MB_ADDR2_SAME" runat="server" Text="同上" onclick="doSame();" /><br />
                            <asp:DropDownList ID="MB_CITY1" runat="server">
                            </asp:DropDownList>
                            市
                            <asp:DropDownList ID="MB_VLG1" runat="server" onChange="MB_VLG1_onChange();"  >
                                <asp:ListItem Text="《鄉鎮市區》" Value=""></asp:ListItem>
                            </asp:DropDownList>
                            區
                            <asp:TextBox ID="MB_ADDR2" runat="server" CssClass="bordertxt" width="60%" onChange="MB_ADDR2_onChange();" />
                            郵遞區號 :
                            <asp:TextBox ID="TXT_MB_ZIP1" runat="server" Columns="3" CssClass="bordernum" />
                        </td>
                    </tr>
                    <tr id="TR_MB_FAMILY" runat="server" visible="false">
                        <td width="15%" class="th1c_b">
                            會員類別
                        </td>
                        <td width="85%" class="td2_b" colspan="3">
                            <asp:RadioButtonList ID="MB_FAMILY" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow" />
                        </td>
                    </tr>
                    <tr>
                        <td width="15%" class="th1c_b">
                            護持會員
                        </td>
                        <td width="85%" class="td2_b" colspan="3">
                            <asp:CheckBoxList ID="VIP" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow" >
                                <asp:ListItem Value="1" Text="長期護持會員" />
                                <asp:ListItem Value="2" Text="種子護法" />
                            </asp:CheckBoxList>
                        </td>
                    </tr>
                    <tr id="TR_FAMILY" runat="server" visible="false">
                        <td width="15%" class="th1c_b">
                            家族人員
                        </td>
                        <td width="85%" class="td2_b" colspan="3">
                            <asp:DataList ID="RP_MB_FAMILY" runat="server" RepeatDirection="Vertical" RepeatLayout="Table" >
                                <ItemStyle Font-Size="11pt" />
                                <ItemTemplate>
                                    <asp:Literal ID="LTL_MB_FAMSEQ" runat="server" />
                                    <asp:Literal ID="LTL_MB_NAME" runat="server" />
                                    <asp:ImageButton ID="IMG_DEL" runat="server" CommandName="DEL" CommandArgument='<%#Container.DataItem("MB_FAMSEQ")%>' />
                                </ItemTemplate>
                            </asp:DataList>
                            <asp:Button ID="btnChoose" runat="server" CssClass="bt" Text="挑選" />
                            <asp:PlaceHolder ID="PLH_F_CHOOSE" runat="server" Visible="false" >
                                <table class="CRTable_Top" width="100%" cellspacing="0">
                                    <asp:PlaceHolder ID="PLH_F_QRY_PARAS" runat="server" >
                                    <tr>
                                        <td width="100%" class="th1_b" colspan="4" >
                                            請先輸入
                                        </td>
                                    </tr>
                                    <tr>
                                        <td width="100%" class="th1_b" colspan="4" >
                                            <asp:RadioButton ID="RB_F_QRY_1" runat="server" GroupName="RB_F_QRY" />
                                            手機:
                                            <asp:TextBox ID="txt_F_QRY_Mobile" runat="server" CssClass="bordernum" style="text-align:left" />
                                            <asp:RadioButton ID="RB_F_QRY_2" runat="server" GroupName="RB_F_QRY" />
                                            電話:
                                            <asp:TextBox ID="txt_F_QRY_Phone_Pre" runat="server" CssClass="bordernum" Columns="2" style="text-align:left" />
                                            ─
                                            <asp:TextBox ID="txt_F_QRY_Phone" runat="server" CssClass="bordernum" style="text-align:left" />
                                            <asp:RadioButton ID="RB_F_QRY_3" runat="server" GroupName="RB_F_QRY" />
                                            姓名:
                                            <asp:TextBox ID="txt_F_QRY_MB_NAME" runat="server" CssClass="bordertxt" style="text-align:left" />
                                            &nbsp;&nbsp;
                                            <asp:Button ID="bt_F_Qry" runat="server" CssClass="bt" Text="確　　定" />
                                        </td>
                                    </tr>
                                    </asp:PlaceHolder>
                                    <asp:PlaceHolder ID="PLH_F_QRY_SAME" runat="server" Visible="false" >
                                    <tr>
                                        <td width="10%" class="td2c_b">
                                            點選
                                        </td>
                                        <td width="25%" class="td2c_b">
                                            姓名
                                        </td> 
                                        <td width="55%" class="td2_b">
                                            通訊地址
                                        </td>                
                                        <td width="10%" class="td2c_b">
                                            會員編號
                                        </td>                                               
                                    </tr>                    
                                    <asp:Repeater ID="RP_F_QRY_SAME" runat="server" >
                                        <ItemTemplate>
                                            <tr>
                                                <td width="10%" class="td2c_b">
                                                    <!--點選-->
                                                    <asp:CheckBox ID="CKB_CHOOSE" runat="server" />
                                                    <input type="hidden" id="HID_MB_MEMSEQ" runat="server" value='<%#Container.DataItem("MB_MEMSEQ")%>' />
                                                 </td>
                                                <td width="25%" class="td2c_b">
                                                    <!--姓名-->
                                                    <asp:LinkButton ID="MB_NAME" runat="server" style="text-decoration:underline;font-weight:bold" CommandArgument='<%#Container.DataItem("MB_MEMSEQ")%>' Text='<%#Container.DataItem("MB_NAME")%>'></asp:LinkButton>
                                                </td> 
                                                <td width="55%" class="td2_b">
                                                    <!--通訊地址-->
                                                    <%# Container.DataItem("MB_CITY")%><%# Container.DataItem("MB_VLG")%><%#Container.DataItem("MB_ADDR")%>
                                                </td>                
                                                <td width="10%" class="td2c_b">
                                                    <!--會員編號-->
                                                    <%# getMB_MEMSEQ(Container.DataItem("MB_MEMSEQ"), Container.DataItem("MB_AREA"))%>
                                                </td>                                               
                                            </tr>                    
                                        </ItemTemplate>
                                    </asp:Repeater>
                                    <tr>
                                        <td width="100%" class="th1_b" colspan="4" >
                                            <asp:Button ID="btn_F_YES" runat="server" CssClass="bt" Text="確定" />
                                            <asp:Button ID="btn_F_ReChoose" runat="server" CssClass="bt" Text="重新挑選" />
                                        </td>
                                    </tr>
                                    </asp:PlaceHolder>
                                </table>
                            </asp:PlaceHolder>
                        </td>
                    </tr>
                </table>
                <asp:PlaceHolder ID="PLH_BUTTON" runat="server">
                <div align="center">
                    <asp:Button ID="bt_Back" runat="server" CssClass="bt" Text="回選擇畫面" />
                    &nbsp;
                    <asp:Button ID="bt_Save" runat="server" CssClass="bt" Text="儲　　存" />
                    &nbsp;
                    <asp:Button ID="bt_ProcMB_MEMREV" runat="server" CssClass="bt" Text="繼續繳款作業" Visible=false />
                </div>
                </asp:PlaceHolder>
            </asp:PlaceHolder>

            <asp:PlaceHolder ID="PLH_MB_MEMREV" runat="server" Visible=false >
                <table class="CRTable_Top" width="100%" cellspacing="0">
                    <tr>
                        <td width="15%" class="th1c_b">
                            法名/姓名
                        </td>
                        <td width="35%" class="td2_b">
                            <Asp:Label ID="LTL_MEMREV_MB_NAME" runat="server" />
                        </td>
                        <td width="15%" class="th1c_b">
                            會員編號
                        </td>
                        <td width="35%" class="td2_b">
                            <Asp:Literal ID="LTL_MEMREV_MB_MEMSEQ" runat="server" />
                            <input type="hidden" id="HID_MB_SEQNO" runat="server" />
                        </td>
                    </tr>
                    <tr>
                        <td width="15%" class="th1c_b">
                            通訊地址
                        </td>
                        <td width="85%" class="td2_b" colspan=3 >
                            <Asp:Literal ID="LTL_MEMREV_MB_ADDR" runat="server" />
                        </td>
                    </tr>
                    <tr>
                        <td width="15%" class="th1c_b">
                            繳款日期
                        </td>
                        <td width="35%" class="td2_b" >
                            <asp:TextBox ID="MB_TX_DATE_YYY" runat="server" CssClass="bordernum" Columns=4 MaxLength=4 />
                            年
                            <asp:TextBox ID="MB_TX_DATE_MM" runat="server" CssClass="bordernum" Columns=2 MaxLength=2 />
                            月
                            <asp:TextBox ID="MB_TX_DATE_DD" runat="server" CssClass="bordernum" Columns=2 MaxLength=2 />
                            日
                        </td>
                        <td width="15%" class="th1c_b">
                            繳款方式
                        </td>
                        <td width="35%" class="td2_b">
                            <asp:RadioButtonList ID="MB_FEETYPE" runat="server" RepeatDirection=Horizontal RepeatLayout=Flow />
                            <asp:Label ID="LBL_ONLY" runat="server" Text="隨緣" style="display:none" />
                        </td>
                    </tr>
                    <tr>
                        <td width="15%" class="th1c_b">
                            功德項目
                        </td>
                        <td width="35%" class="td2_b">
                            <asp:DropDownList ID="MB_ITEMID" runat="server" onchange="MB_ITEMID_onchange(this);" />
                            <span id="PLH_MB_FAMILY" runat="server" style="display:none" >
                                會員類別：
                                <asp:RadioButtonList ID="RBL_MEMREV_MB_FAMILY" runat="server" RepeatDirection=Horizontal RepeatLayout=Flow AutoPostBack="true" />
                            </span>
                            <span id="PLH_PROJCODE" runat="server" style="display:none" >
                                專案名稱：
                                <asp:DropDownList ID="PROJCODE" runat="server" />
                            </span>                    
                        </td>
                        <td width="15%" class="th1c_b">
                            繳款金額
                        </td>
                        <td width="35%" class="td2_b">
                            <asp:TextBox ID="MB_TOTFEE" runat="server" Columns="7" MaxLength=9 onchange="MB_TOTFEE_onchange(this);" />元
                            <span id="PLH_MB_TOTFEE_MM" runat="server" style="display:none" >
                                每月金額：
                                <asp:TextBox ID="MB_TOTFEE_MM" runat="server" CssClass="bordernum" Columns="7" MaxLength=9 />元
                            </span>
                            <span id="PLH_VIP_ACCAMT" runat="server" style="display:none" >
                                累計繳款：
                                <asp:TextBox ID="ACCAMT" runat="server" CssClass="nobordernum" Columns="7" MaxLength=9 />元
                                <input type="hidden" id="HID_ACCAMT" runat="server" />
                            </span>
                        </td>
                    </tr>
                    <tr id="TR_FUND" runat="server" style="display:none" >
                        <td width="15%" class="th1c_b">
                            繳款分配項目
                        </td>
                        <td width="35%" class="td2_b" colspan="3">
                            <asp:CheckBox ID="FUND1" runat="server" Text="護僧基金" />
                            <asp:CheckBox ID="FUND2" runat="server" Text="建設基金" />
                            <asp:CheckBox ID="FUND3" runat="server" Text="弘法基金" />
                        </td>
                    </tr>
                    <tr>
                        <td width="15%" class="th1c_b">
                            收據捐款名稱
                        </td>
                        <td width="35%" class="td2_b">
                            <asp:CheckBox ID="CB_MB_RECNAME" runat="server" onclick="CB_MB_RECNAME_onclick(this)" />同繳款人
                            <BR />
                            <asp:TextBox ID="MB_RECNAME" runat="server" CssClass="bordertxt" />
                        </td>
                        <td width="15%" class="th1c_b">
                            所屬區
                        </td>
                        <td width="35%" class="td2_b">
                            <asp:Literal ID="LTL_MEMREV_MB_AREA" runat="server" />
                            &nbsp;
                            委員：
                            <asp:Literal ID="LTL_MEMREV_MB_LEADER" runat="server" />
                        </td>
                    </tr>
                    <tr>
                        <td width="15%" class="th1c_b">
                            <asp:label ID="LTL_MB_ITEMID_TYP" runat="server" Text="會費期間" />
                        </td>
                        <td width="35%" class="td2_b">
                            <span ID="PLH_MB_MEMFEE" runat="server"  >
                                <asp:TextBox ID="MB_MEMFEE_SY" runat="server" CssClass="bordernum" Columns=4 MaxLength=4 onblur="MB_MEMFEE_SYM_CHANGE();" />
                                年
                                <asp:TextBox ID="MB_MEMFEE_SM" runat="server" CssClass="bordernum" Columns=2 MaxLength=2 onblur="MB_MEMFEE_SYM_CHANGE();" />
                                月至
                                <asp:TextBox ID="MB_MEMFEE_EY" runat="server" CssClass="bordernum" Columns=4 MaxLength=4 />
                                年
                                <asp:TextBox ID="MB_MEMFEE_EM" runat="server" CssClass="bordernum" Columns=2 MaxLength=2 />
                                月                    
                            </span>
                            <asp:TextBox ID="MB_DESC" runat="server" CssClass="bordertxt" style="display:none" />
                        </td>
                        <td width="15%" class="th1c_b">
                            付款方式
                        </td>
                        <td width="35%" class="td2_b">
                            <asp:RadioButtonList ID="MB_PAY_TYPE" runat="server" RepeatDirection=Horizontal RepeatLayout=Flow onclick="MB_PAY_TYPE_onclick();" />
                            <div id="PLH_MB_PAY_TYPE_N" runat="server" style="display:none" >
                                票據到期日:
                                <asp:TextBox ID="NOTE_DUE_DATE_YYY" runat="server" CssClass="bordernum" MaxLength=4 Columns=4 />
                                年
                                <asp:TextBox ID="NOTE_DUE_DATE_MM" runat="server" CssClass="bordernum" MaxLength=2 Columns=2 />
                                月
                                <asp:TextBox ID="NOTE_DUE_DATE_DD" runat="server" CssClass="bordernum" MaxLength=2 Columns=2 />
                                日
                                <BR/>
                                票據號碼:
                                <asp:TextBox ID="NOTE_NO" runat="server" CssClass="bordertxt" MaxLength=20 />
                                <BR/>
                                發票行:
                                <asp:TextBox ID="NOTE_BANK" runat="server" CssClass="bordertxt" MaxLength=20 />
                                銀行
                                <asp:TextBox ID="NOTE_BR" runat="server" CssClass="bordertxt" MaxLength=20 />
                                分行
                                <BR/>
                                發票人:
                                <asp:TextBox ID="NOTE_HOLDER" runat="server" CssClass="bordertxt" MaxLength=50 />
                                <BR/>
                                票據金額:
                                <asp:TextBox ID="NOTE_AMT" runat="server" CssClass="bordernum" MaxLength=9 />
                            </div>
                        </td>
                    </tr>
                    <tr>
                        <td width="15%" class="th1c_b">
                            是否開立收據
                        </td>
                        <td id="TD_PRINT" runat="server" colspan="3" width="35%" class="td2_b">
                            <asp:RadioButtonList ID="RBL_MB_SEND_PRINT" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow" AutoPostBack="true" >
                                <asp:ListItem Text="是" Value="1" />
                                <asp:ListItem Text="否" Value="0" />
                            </asp:RadioButtonList>
                        </td>
                        <asp:PlaceHolder ID="PLH_MB_SEND_PRINT" runat="server" Visible="false" >
                        <td width="15%" class="th1c_b">
                            給收據方式
                        </td>
                        <td width="35%" class="td2_b">
                            <asp:DropDownList ID="DDL_MB_SEND_PRINT" runat="server" />
                        </td>
                        </asp:PlaceHolder>
                    </tr>
                </table>
                <div align="center">
                    <asp:Button ID="bt_MB_MEMREV_CANCEL" runat="server" Text="取　　消" CssClass="bt"  Visible=false/>
                    <asp:Button ID="bt_MB_MEMREV_YES" runat="server" Text="確　　定" CssClass="bt" Visible=false  />
                    <asp:Button ID="bt_MB_MEMREV_Add" runat="server" Text="新　　增" CssClass="bt" />
                </div>
                <asp:PlaceHolder ID="PLH_Repeater" runat="server" >
                <table class="CRTable_Top" width="100%" cellspacing="0" >
                    <tr>
                        <td width="20%" class="th1c_b">
                            繳款日期
                        </td>
                        <td width="10%" class="th1c_b">
                            功德項目
                        </td>
                        <td width="15%" class="th1c_b">
                            繳款金額
                        </td>
                        <td width="15%" class="th1c_b">
                            收據捐款名稱
                        </td>
                        <td width="10%" class="th1c_b">
                            所屬區
                        </td>
                        <td width="10%" class="th1c_b">
                            委員
                        </td>
                        <td width="10%" class="th1c_b">
                            修改
                        </td>
                        <td width="10%" class="th1c_b">
                            刪除
                        </td>
                    </tr>
                    <asp:Repeater ID="RP_MB_MEMREV" runat="server" >
                        <ItemTemplate>
                            <tr>
                                <td width="20%" class="td2c_b">
                                    <!--繳款日期-->
                                    <asp:Literal ID="LTL_MB_TX_DATE" runat="server" Text='<%#getROCDate(Container.DataItem("MB_TX_DATE"))%>' />
                                </td>
                                <td width="10%" class="td2c_b">
                                    <!--功德項目-->
                                    <asp:Literal ID="LTL_MB_ITEMID" runat="server" Text='<%#getMB_ITEMID(Container.DataItem("MB_ITEMID"))%>' />
                                </td>
                                <td width="15%" class="td2r_b">
                                    <!--繳款金額-->
                                    <asp:Literal ID="LTL_MB_TOTFEE" runat="server" Text='<%#getMB_TOTFEE(Container.DataItem("MB_TOTFEE"))%>' />
                                </td>
                                <td width="15%" class="td2c_b">
                                    <!--收據捐款名稱-->
                                    <%#Container.DataItem("MB_RECNAME")%>
                                </td>
                                <td width="10%" class="td2c_b">
                                    <!--所屬區-->
                                    <asp:Literal ID="LTL_MB_AREA" runat="server" />
                                </td>
                                <td width="10%" class="td2c_b">
                                    <!--委員-->
                                    <asp:Literal ID="LTL_MB_LEADER" runat="server" />
                                </td>
                                <td width="10%" class="td2c_b">
                                    <!--修改-->
                                    <asp:ImageButton ID="IMG_EDIT" runat="server" CommandName="EDIT" CommandArgument='<%#Container.DataItem("MB_SEQNO")%>' ImageUrl="~/img/edit.gif" />
                                </td>
                                <td width="10%" class="td2c_b">
                                    <!--刪除-->
                                    <asp:ImageButton ID="IMG_DEL" runat="server" CommandName="DELETE" CommandArgument='<%#Container.DataItem("MB_SEQNO")%>' ImageUrl="~/img/trashcan.gif" />
                                </td>
                            </tr>
                        </ItemTemplate>
                    </asp:Repeater>
                </table>
                <div align="center">
                    <asp:Button ID="bt_BackMain" runat="server" CssClass="bt" Text="回選擇畫面" />
                </div>
                </asp:PlaceHolder>
            </asp:PlaceHolder>
            <!--顯示下方外框線-->
            <!--#include virtual="~/inc/MBSCTableEnd.inc"-->
        </td>
    </tr>
</table>
    </form>
</body>
</html>
