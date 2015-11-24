<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="PagerMenu.ascx.vb"
    Inherits="MBSC.UICtl.PagerMenu" %>
<table id="tbPagingNavigation" runat="server" cellpadding="0" cellspacing="0" width="98%"
    border="0">
    <tr>
        <td align="left" style="white-space: nowrap; height: 8%; font-size: 13px; font-style: normal;
            color: #000000; padding-top: 3px; padding-bottom: 3px">
            目前在第<asp:Label ID="lblPageIndex" runat="server" Text="0"></asp:Label>頁<span id="span1"
                runat="server"> <span id="span2" runat="server">|</span></span>共
            <asp:Label ID="lblTotalPage" runat="server" Text="0"></asp:Label>頁(每頁
            <asp:Label ID="lblTotalRecord" runat="server" Text="0"></asp:Label>筆)
        </td>
        <td align="right" style="height: 24px; font-size: 13px; font-style: normal; color: #000000;
            padding-top: 3px; padding-bottom: 3px">
            <asp:ImageButton ID="imgBtnFirst" runat="server" ImageUrl="~/img/imgFirstPage.gif" />
            <asp:LinkButton ID="btnNavFirst" ForeColor="Blue" runat="server" CausesValidation="False"
                CommandName="FIRST" Text="第一頁" OnClick="NavigationButtonClick"></asp:LinkButton>
            <span id="spanFirst" runat="server">|</span>
            <asp:ImageButton ID="imgBtnPrevious" runat="server" ImageUrl="~/img/imgBackPage.gif" />
            <asp:LinkButton ID="btnNavPrevious" ForeColor="Blue" runat="server" CausesValidation="False"
                CommandName="PREVIOUS" Text="上一頁" OnClick="NavigationButtonClick"></asp:LinkButton>
            <span id="spanSecond" runat="server">|</span>
            <asp:LinkButton ID="btnNavNext" ForeColor="Blue" runat="server" CausesValidation="False"
                CommandName="NEXT" Text="下一頁" OnClick="NavigationButtonClick"></asp:LinkButton><span
                    id="spanThird" runat="server">|</span>
            <asp:ImageButton ID="imgBtnNext" runat="server" ImageUrl="~/img/imgNextPage.gif" />
            <asp:LinkButton ID="btnNavLast" ForeColor="Blue" runat="server" CausesValidation="False"
                CommandName="LAST" Text="最末頁" OnClick="NavigationButtonClick"></asp:LinkButton>
            <asp:ImageButton ID="imgBtnLast" runat="server" ImageUrl="~/img/imgLastPage.gif" />
            <asp:DropDownList ID="ddlJumpPage" runat="server" CssClass="fontSize" AutoPostBack="true">
            </asp:DropDownList>
        </td>
    </tr>
</table>
