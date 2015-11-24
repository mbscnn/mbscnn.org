<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SY_TOOLBAR.aspx.vb" Inherits="MBSC.SY_TOOLBAR" %>

<script language="VB" runat="server">
</script>
<html>
<head>
    <title>上方功能列</title>
    <meta http-equiv="Content-Type" content="text/html; charset=big5">
    <link rel="stylesheet" href="<%=com.Azion.EloanUtility.UIUtility.getRootPath()%>/css/homeword.css"
        type="text/css">
    <Script lang="JavaScript" src="<%=com.Azion.EloanUtility.UIUtility.getRootPath() + "/js/jquery.js"%>"></Script>
    <Script lang="JavaScript" src="<%=com.Azion.EloanUtility.UIUtility.getRootPath() + "/js/jquery.validate.js"%>"></Script>
    <Script lang="JavaScript" src="<%=com.Azion.EloanUtility.UIUtility.getRootPath() + "/js/jquery.validate.extension.js"%>"></Script>
    <Script lang="JavaScript" src="<%=com.Azion.EloanUtility.UIUtility.getRootPath() + "/js/Calendar.js"%>"></Script>
    <Script lang="JavaScript" src="<%=com.Azion.EloanUtility.UIUtility.getRootPath() + "/js/JSUtil.js"%>"></Script>
    
    <script type="text/javascript" language="JavaScript">
<!--
        // 圖片onmouseout
        function MM_swapImgRestore() { //v3.0
            var i, x, a = document.MM_sr; for (i = 0; a && i < a.length && (x = a[i]) && x.oSrc; i++) x.src = x.oSrc;
        }

        // 加載圖片
        function MM_preloadImages() { //v3.0
            var d = document; if (d.images) {
                if (!d.MM_p) d.MM_p = new Array();
                var i, j = d.MM_p.length, a = MM_preloadImages.arguments; for (i = 0; i < a.length; i++)
                    if (a[i].indexOf("#") != 0) { d.MM_p[j] = new Image; d.MM_p[j++].src = a[i]; }
            }
        }

        // 找控件
        function MM_findObj(n, d) { //v4.0
            var p, i, x; if (!d) d = document; if ((p = n.indexOf("?")) > 0 && parent.frames.length) {
                d = parent.frames[n.substring(p + 1)].document; n = n.substring(0, p);
            }
            if (!(x = d[n]) && d.all) x = d.all[n]; for (i = 0; !x && i < d.forms.length; i++) x = d.forms[i][n];
            for (i = 0; !x && d.layers && i < d.layers.length; i++) x = MM_findObj(n, d.layers[i].document);
            if (!x && document.getElementById) x = document.getElementById(n); return x;
        }

        // 圖片onMouseOver
        function MM_swapImage() { //v3.0
            var i, j = 0, x, a = MM_swapImage.arguments; document.MM_sr = new Array; for (i = 0; i < (a.length - 2); i += 3)
                if ((x = MM_findObj(a[i])) != null) { document.MM_sr[j++] = x; if (!x.oSrc) x.oSrc = x.src; x.src = a[i + 2]; }
        }

        // 開窗
        function open_form(sURL, sW, sH) {
            if (sW == '') sW = 0;
            if (sH == '') sW = 0;
            openWindow(sURL, "Open", 0, 0, 1, 0, sW, sH);
        }

        // open的形式開窗
        function openWindow(url, name, rs, mnu, scrl, stat, w, h) {
            var resize = (rs) ? "resizable," : "";
            var menu = (mnu) ? "menubar," : "";
            var status = (stat) ? "status," : "";
            var scroll = (scrl) ? "scrollbars" : "";
            var screenW, screenH, left, top;

            // Get the screen dimensions.
            if (screen) { screenW = screen.width; screenH = screen.height; }
            else if (window.screen) { screenW = window.screen.width; screenH = window.screen.height; }

            if (w == 0) w = 800;
            if (h == 0) h = 600;

            // Center the window on the screen.
            left = ((screenW / 2) - (w / 2));
            top = ((screenH / 3) - (h / 2));
            if (top < 5) top = 5;

            // Establish the new window's features.
            var features = "width=" + w + ",height=" + h + ",left=" + left + ",top=" + top + "," + status + resize + menu + scroll;

            // Show the window.
            popupWin = window.open(url, name, features);
        }

        // 開啟新窗體
        function showmodal(url, ww, wh) {
            showModalDialog(url, window, "status=no;dialogWidth:" + ww + "px;dialogHeight:" + wh + "px;");
        }
   
        // 重新刷新頁面
        function do_reload(i) {
            if (i == 0) {
                parent.frames[1].frames[0].location.reload();
            } else if (i == 1)
            {
                parent.location.reload();
                //parent.frames[1].frames[1].location.reload();
                //parent.frames[1].frames[1].location.assign(parent.frames[1].frames[1].location);
            } else {
                parent.frames[1].frames[0].location.reload();
                parent.frames[1].frames[1].location.reload();
            }
        }

        // 檢核已處理工作和待處理工作是否有選擇一筆案件
        // 沒有已處理及待處理工作清單時，直接跳出
        function checkCheckedItemNum() {

            var $caseTable = null;
            var checkedCaseNum = 0;
            var caseId = "";

            //  待處理案件清單
            var caseList = parent.frames["downFrame"].frames["mainFrame"].document.all["dgCase"];

            // 已處理工作清單
            var completeCaseList = parent.frames["downFrame"].frames["mainFrame"].document.all["dgCompleteCase"];

            if (caseList != undefined) {
                $caseTable = $(caseList.parentElement);
            } else if (completeCaseList != undefined) {
                $caseTable = $(completeCaseList.parentElement);
            } else {
                return false;
            }

            checkedCaseNum = $caseTable.find(":input:radio:checked").length

            if (checkedCaseNum == 0)
            {
                alert("請選擇一筆案件資料！");

                return false;
            }
            else {
                caseId = jQuery.trim($caseTable.find(":input:radio:checked").parent().parent().text()).split(" ")[0];
                var url;
                url = 'SY_FLOWSTEP.aspx?caseid=' + caseId
                showmodal(url, '600', '400');
                parent.location.reload();
            }

            return true;
        }

        // 佇列取件
        function do_viewqueue() {
            //'showModalDialog(".aspx", window, "status=no;dialogWidth:800px;dialogHeight:400px");
              showmodal('SY_CASEQUEUE.aspx', '800', '400');
              parent.location.reload();
        }

        // 個人資訊開窗
        function Show_Info() {
            showmodal('SY_USERINFO.aspx?userid=<%=Session("UserID")%>', '290', '310');
            parent.location.reload();
        }
        
        // 切換身份
        function Show_ChangeUser()
        {
            showmodal('SY_CHANGEUSER.aspx', '800', '300');
            parent.location.reload();
        }
-->
    </script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0"
    marginheight="0" onload="MM_preloadImages('<%=com.Azion.EloanUtility.UIUtility.getImgPath()%>/top_icon001a.jpg','<%=com.Azion.EloanUtility.UIUtility.getImgPath()%>/top_icon002a.jpg','<%=com.Azion.EloanUtility.UIUtility.getImgPath()%>/top_icon003a.jpg','<%=com.Azion.EloanUtility.UIUtility.getImgPath()%>/top_icon004a.jpg','<%=com.Azion.EloanUtility.UIUtility.getImgPath()%>/top_icon005a.jpg','<%=com.Azion.EloanUtility.UIUtility.getImgPath()%>/top_icon006a.jpg','<%=com.Azion.EloanUtility.UIUtility.getImgPath()%>/top_icon007a.jpg')">
    <div id="outError" runat="server" />
    <form name="toolbar">
    <table border="0" cellspacing="0" cellpadding="0" width="100%" bordercolor="#FFFFCC">
        <tr>
            <td width="100%" valign="top" height="81">
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr background="<%=com.Azion.EloanUtility.UIUtility.getImgPath()%>/top_t01.gif">
                        <td background="<%=com.Azion.EloanUtility.UIUtility.getImgPath()%>/top_back.jpg"
                            width="722">
                            <img src="<%=com.Azion.EloanUtility.UIUtility.getImgPath()%>/top_tital.gif" width="209"
                                height="45">
                        </td>
                    </tr>
                    <tr style="background-image: url(<%=com.Azion.EloanUtility.UIUtility.getImgPath()%>/top_barblue.jpg);">
                        <td width="100%">
                            <div id="toolbar1_non_remote_div" style="display: block">
                                <table width="100%" border="0" cellspacing="0" cellpadding="0" background="<%=com.Azion.EloanUtility.UIUtility.getImgPath()%>/top_barblue.jpg">
                                    <tr>
                                        <td width="150">
                                            <img src="<%=com.Azion.EloanUtility.UIUtility.getImgPath()%>/spacer.gif" width="150"
                                                height="20">
                                        </td>
                                        <td width="90">
                                            <a href="javascript:Show_Info();" alt="個人資訊" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('Image1','','<%=com.Azion.EloanUtility.UIUtility.getImgPath()%>/top_icon001a.jpg',1)">
                                                <img src="<%=com.Azion.EloanUtility.UIUtility.getImgPath()%>/top_icon001.jpg" border="0"
                                                    name="Image1"></a>
                                        </td>
                                        <td width="90" style="display: block">
                                            <a href="javascript:Show_ChangeUser();"  alt="切換身分" onmouseout="MM_swapImgRestore()"
                                                onmouseover="MM_swapImage('Image2','','<%=com.Azion.EloanUtility.UIUtility.getImgPath()%>/top_icon0021a.jpg',1)">
                                                <img src="<%=com.Azion.EloanUtility.UIUtility.getImgPath()%>/top_icon0021.jpg" border="0"
                                                    name="Image2"></a>
                                        </td>
                                        <td width="90">
                                            <a  alt="案件狀態" onmouseout="MM_swapImgRestore()"
                                                onmouseover="MM_swapImage('Image3','','<%=com.Azion.EloanUtility.UIUtility.getImgPath()%>/top_icon003a.jpg',1)" onclick="return checkCheckedItemNum();">
                                                <img src="<%=com.Azion.EloanUtility.UIUtility.getImgPath()%>/top_icon003.jpg" border="0"
                                                    name="Image3"></a>
                                        </td>
                                        <td width="90" style="display: block">
                                            <a href="javascript:do_viewqueue();"
                                                alt="佇列取件" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('Image4','','<%=com.Azion.EloanUtility.UIUtility.getImgPath()%>/top_icon0041a.jpg',1)">
                                                <img src="<%=com.Azion.EloanUtility.UIUtility.getImgPath()%>/top_icon0041.jpg"  border="0"
                                                    name="Image4"></a>
                                        </td>
                                        <td width="90">
                                            <a href="javascript:do_reload(1)" alt="重新整理" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('Image6','','<%=com.Azion.EloanUtility.UIUtility.getImgPath()%>/top_icon006a.jpg',1)">
                                                <img src="<%=com.Azion.EloanUtility.UIUtility.getImgPath()%>/top_icon006.jpg" id="imgRefresh" width="84"
                                                    height="20" border="0" name="Image6" onclick ="do_reload('1')"></a>
                                        </td>
                                        <td width="100%">
                                            &nbsp;<font color="yellow" size="4"><span id="sp_marquee" runat="server"></span></font>
                                        </td>
                                    </tr>
                                </table>
                            </div>
                        </td>
                    </tr>
                    <tr bgcolor="#ACD8FF">
                        <td width="100%">
                            <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                <tr>
                                    <td class="homeword" width="11">
                                        &nbsp;
                                    </td>
                                    <td class="homeword" width="1">
                                        &nbsp;
                                    </td>
                                    <td width="26">
                                        <img src="<%=com.Azion.EloanUtility.UIUtility.getImgPath()%>/top_bar.jpg" width="26"
                                            align="top" height="17" hspace="0" vspace="0">
                                    </td>
                                    <td class="smallbar" width="18">
                                        <img src="<%=com.Azion.EloanUtility.UIUtility.getImgPath()%>/top_blueline.gif" width="18"
                                            height="17" align="baseline" hspace="0" vspace="0">
                                    </td>
                                    <%=m_SSList.toString()%>
                                    <td class="smallbar" width="33">
                                        &nbsp;
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
