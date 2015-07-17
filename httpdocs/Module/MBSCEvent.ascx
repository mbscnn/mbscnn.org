<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="MBSCEvent.ascx.vb" Inherits="MBSC.UICtl.MBSCEvent" %>
<div id="DIV_EVENT ">
    <ul>
        <asp:Repeater ID="RP_EVENT" runat="server" >
            <ItemTemplate>
                <li style="font-family:細明體;font-size:10pt">
                    <asp:LinkButton ID="LBL_EVENT" runat="server" Text='<%#Container.DataItem("EVENT")%>' CommandArgument='<%#Container.DataItem("CRETIME")%>' CommandName="CRETIME" />
                </li>
            </ItemTemplate>
        </asp:Repeater>
    </ul>
</div>