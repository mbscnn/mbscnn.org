<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="VTab.ascx.vb" Inherits="MBSC.UICtl.VTab" %>
		
    <ul class="nav navbar-nav navbar-right">
                                    <asp:Repeater runat="server" ID="RP_Tab">
                                        <ItemTemplate>
                                            <li id="Tab_Li" runat="server">
                                                <span ID="HID_CODEID" runat="server" visible="false" ><%#Container.DataItem("CODEID")%></span>                                              
                                                <asp:HyperLink ID="HPL_LV1" runat="server" data-toggle="dropdown" class="dropdown-toggle" Style="font-size:11pt" Text='<%#Container.DataItem("TEXT")%>' NavigateUrl="#" />
                                            
                                            
                                            <ul class="dropdown-menu">
                                            <asp:PlaceHolder ID="PLH_LV2" runat="server" Visible="true">
                                            <asp:Repeater runat="server" ID="RP_Tab_LV2" OnItemDataBound="RP_Tab_LV2_ItemDataBound">
                                                <ItemTemplate>                                              
                                                    <!-- if has children -->                                                    
                                           
                                                    <li id="Tab_Li" runat="server" class="dropdown dropdown-submenu" >                                                      
                                                            <span ID="HID_CODEID" runat="server" visible="false" ><%#Container.DataItem("CODEID")%></span>                                                          
                                                            <asp:HyperLink ID="HPL_LV2" runat="server" class="dropdown-toggle" Style="font-size:11pt" data-toggle="dropdown" Text='<%#Container.DataItem("TEXT")%>' NavigateUrl="#" />
                                                                                                                 
                                                        <ul class="dropdown-menu">
                                                        <asp:PlaceHolder ID="PLH_LV3" runat="server" Visible="true">
                                                        <asp:Repeater runat="server" ID="RP_Tab_LV3" OnItemDataBound="RP_Tab_LV3_ItemDataBound">                                                           
                                                            <ItemTemplate>
                                                                <li id="Tab_Li" runat="server" >
                                                                    <span ID="HID_CODEID" runat="server" visible="false" ><%#Container.DataItem("CODEID")%></span>
                                                                    <asp:HyperLink ID="HPL_LV3" runat="server" Text='<%#Container.DataItem("TEXT")%>' NavigateUrl="#" />
                                                                </li>
                                                            </ItemTemplate>                                                                                                             
                                                        </asp:Repeater>
                                                        </asp:PlaceHolder>
                                                        </ul>                                                       
                                                    </li>
                                                </ItemTemplate>
                                            </asp:Repeater>
                                            </asp:PlaceHolder>
                                            </ul>
                                            </li>
                                        </ItemTemplate>
                                    </asp:Repeater>
                                </ul>
                                <script>
                                    (function ($) {
                                        $(document).ready(function () {
                                            $('ul.dropdown-menu [data-toggle=dropdown]').on('click', function (event) {
                                                event.preventDefault();
                                                event.stopPropagation();
                                                $(this).parent().siblings().removeClass('open');
                                                $(this).parent().toggleClass('open');
                                            });
                                        });
                                    })(jQuery);
        </script>