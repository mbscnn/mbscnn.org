<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SY_CHANGEUSER.aspx.vb"
    Inherits="MBSC.SY_CHANGEUSER" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>切換身分</title>
    <base target ="_self" />
    <!--共用的SY CSS-->
    <!-- #include virtual="~/inc/SYCSS.inc" -->
</head>
<body>
    <form id="form1" runat="server">
    <div>
       
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="false"  GridLines="Both"  CellPadding="1" 
              HeaderStyle-BackColor="#7779AF" HeaderStyle-ForeColor="White" Width="100%"  BorderStyle="Solid" BorderWidth="1">
            <Columns>
                <asp:TemplateField>
                    <ItemTemplate>
                        <asp:RadioButton ID="rdbLoginStaffid" runat="server" GroupName="gp" AutoPostBack="true" OnCheckedChanged="select_RowCommand" />
                    </ItemTemplate>
                </asp:TemplateField>

                <asp:BoundField DataField="LoginBrId" HeaderText="登入<br />分行別" HeaderStyle-Width="8%" HtmlEncode="false" />
                <asp:BoundField DataField="LoginBrName" HeaderText="登入<br />分行名稱"  HeaderStyle-Width="12%" HtmlEncode="false" />
                <asp:BoundField DataField="LoginStaffid" HeaderText="登入<br />員編" HeaderStyle-Width="10%" HtmlEncode="false" />
                <asp:BoundField DataField="LoginUserName" HeaderText="登入<br />姓名"  HeaderStyle-Width="10%" HtmlEncode="false" />
                <asp:BoundField DataField="WorkingBrid" HeaderText="被代理<br />分行別" HeaderStyle-Width="8%" HtmlEncode="false" />
                <asp:BoundField DataField="WorkingBrName" HeaderText="被代理<br />分行名稱"  HeaderStyle-Width="12%" HtmlEncode="false" />
                <asp:BoundField DataField="WorkingStaffid" HeaderText="被代理<br />員編" HeaderStyle-Width="10%" HtmlEncode="false" />
                <asp:BoundField DataField="WorkingUserName" HeaderText="被代理<br />姓名"  HeaderStyle-Width="10%" HtmlEncode="false" />
                <asp:BoundField DataField="STARTTIME" HeaderText="代理<br />起日"  HeaderStyle-Width="10%" HtmlEncode="false" />
                <asp:BoundField DataField="ENDTIME" HeaderText="代理<br />迄日"  HeaderStyle-Width="10%" HtmlEncode="false" />
            </Columns>
        </asp:GridView>


        
       
    </div>
    </form>
</body>
</html>
