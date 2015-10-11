<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="Signin.ascx.vb" Inherits="MBSC.UICtl.Signin" %>

<!-- page title -->
<h2 class="section-heading animated fadeInLeftBig text-center" style="color: #020202;margin-top:0;margin-bottom:0;line-height:0.1">
    <asp:Literal ID="LTL_TAB_TITLE" runat="server" />
    <span style="float:right;vertical-align:top;font-size:16pt">
        <asp:PlaceHolder ID="PLH_LOGOUT" runat="server" Visible="false">
            <span style="vertical-align:top;">
                <span style="font-weight: bold;font-size:16pt;color:white">
                    <asp:Literal ID="LTL_MB_NAME" runat="server" />                
                    <asp:Button ID="btLogout" runat="server" Text="登出" class="btn btn-primary btn-xs gradient" />
                    <asp:Literal ID="LTL_TAB_TITLE_HID" runat="server" Visible="false" />
                </span>
            </span>
        </asp:PlaceHolder>       
    </span>
</h2>
                
<hr class="animated fadeInRightBig" > 



<!-- /.page title -->

<div style="clear: both;"></div>