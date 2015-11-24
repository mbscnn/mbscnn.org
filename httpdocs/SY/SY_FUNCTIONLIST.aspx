<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SY_FUNCTIONLIST.aspx.vb"
    Inherits="MBSC.SY_FUNCTIONLIST" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>leftitem1</title>
    <link href="<%=com.Azion.EloanUtility.UIUtility.getRootPath() + "/css/leftitem.css"%>"
        type="text/css" rel="StyleSheet" media="screen">
</head>
<body width="150" backgroundsrc="<%=com.Azion.EloanUtility.UIUtility.getImgPath()%>leftitem/leftitem_bg.gif"
    leftmargin="7">
    <form id="form1" runat="server">
    <span id="outError" runat="server" />
    <div id="Outline">
        <table width="100%" border="0" cellpadding="0" cellspacing="0">
            <tr>
                <td class="td2" height="30">
                    <img src="<%=com.Azion.EloanUtility.UIUtility.getImgPath()%>person2.gif" align="absbottom">
                    <span id="infomation" runat="server" />
                </td>
            </tr>
            <tr>
                <td>
                    <asp:TreeView ID="treeViewSys" runat="server" ExpandDepth="2">
                        <RootNodeStyle ImageUrl="~/img/map2.gif" />
                        <ParentNodeStyle ImageUrl="~/img/folder.png" />
                        <LeafNodeStyle ImageUrl="~/img/file.gif" />
                    </asp:TreeView>
                </td>
            </tr>
        </table>
    </div>
    </form>
</body>
</html>
