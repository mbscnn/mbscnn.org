<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ZipCode.ascx.vb"
    Inherits="MBSC.UICtl.ZipCode" %>
<asp:dropdownlist id="ddlCity" runat="server" AutoPostBack="true">
    <asp:ListItem Value="">請選擇</asp:ListItem>
</asp:dropdownlist>
<asp:dropdownlist id="ddlTown" runat="server" AutoPostBack="true">
    <asp:ListItem Value="">請選擇</asp:ListItem>
</asp:dropdownlist>
<asp:TextBox ID="txtRoad" runat="server" Width="307px"></asp:TextBox>
&nbsp;<asp:PlaceHolder ID="plhPostCode" runat="server">
郵遞區號
<asp:Literal id="ltlPostCode" Runat="server" />
<asp:TextBox ID="txtPostCode" runat="server" Visible="false" MaxLength="5"></asp:TextBox> 
</asp:PlaceHolder>
<%--<asp:RadioButton ID="rdoSw0" runat="server" Text="自有" GroupName="Sw" TabIndex="55" />
<asp:RadioButton ID="rdoSw1" runat="server" Text="租用" GroupName="Sw" TabIndex="60" />--%>