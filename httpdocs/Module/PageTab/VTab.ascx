<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="VTab.ascx.vb" Inherits="MBSC.UICtl.VTab" %>
		<style>
            @import url(http://mbscnn.org/css/font-awesome.min.css);
			* {margin: 0; padding: 0;}

			#accordian {
				background: #004050;
				width: 180px;
				margin: 0px auto 0 auto;
				color: white;
				box-shadow: 
				0 5px 15px 1px rgba(0, 0, 0, 0.6), 
				0 0 200px 1px rgba(255, 255, 255, 0.5);
                float: left;
			}

            #accordian h3 {
              background: #003040;
              background: linear-gradient(#003040, #002535);
            }

			#accordian h3 a {
				padding: 0 10px;
				font-size: 12pt;
                font-family:"微軟正黑體","新細明體";
				line-height: 34px;
				display: block;
				color: white;
				text-decoration: none;
			}
			
			#accordian h3:hover {
				text-shadow: 0 0 1px rgba(255, 255, 255, 0.7);
			}
			
			#accordian h3 span {
				font-size: 12pt;
				margin-right: 10px;
			}
			
			#accordian li {
				list-style-type: none;
			}
			
			#accordian ul ul li a, #accordian h4 {
				color: white;
				text-decoration: none;
				font-size: 11pt;
                font-family:"微軟正黑體","新細明體";
				line-height: 27px;
				display: block;
				padding: 0 25px;

				transition: all 0.15s;
				position: relative;
			}
			
			#accordian ul ul li a:hover {
				background: #003545;
				border-left: 5px solid lightgreen;
			}
			
			#accordian ul ul {
				display: none;
			}
			#accordian li.active>ul {
				display: block;
			}

			#accordian ul ul ul{
				margin-left: 15px; border-left: 1px dotted rgba(0, 0, 0, 0.5);
			}

			#accordian a:not(:only-child):after {
				content: "\f104";
				font-family: fontawesome;
				position: absolute; right: 10px; top: 0;
				font-size: 12pt;
			}
			#accordian .active>a:not(:only-child):after {
				content: "\f107";
			}

            .fixedTop {
                margin-left: 0px !important;
                position:fixed;
                top:0;
                display:block;
                /*overflow-y:scroll;*/
                overflow-y:auto;
                max-height: 100%;
                overflow-x:hidden;
            }
		</style>

        <div id="accordian" >
            <ul>
                <asp:Repeater ID="RP_Tab" runat="server" >
                    <ItemTemplate>
                        <li id="Tab_Li" runat="server">
                            <span ID="HID_CODEID" runat="server" visible="false" ><%#Container.DataItem("CODEID")%></span>
                            <h3><asp:HyperLink ID="HPL_LV1" runat="server" Text='<%#Container.DataItem("TEXT")%>' NavigateUrl="#" /></h3>
                            <asp:PlaceHolder ID="PLH_LV2" runat="server" Visible="false">
                                <ul>
                                    <asp:Repeater ID="RP_Tab_LV2" runat="server" OnItemDataBound="RP_Tab_LV2_ItemDataBound" >
                                        <ItemTemplate>
                                            <li id="Tab_Li" runat="server" >
                                                <span ID="HID_CODEID" runat="server" visible="false" ><%#Container.DataItem("CODEID")%></span>
                                                <asp:HyperLink ID="HPL_LV2" runat="server" Text='<%#Container.DataItem("TEXT")%>' NavigateUrl="#" />
                                                <asp:PlaceHolder ID="PLH_LV3" runat="server" Visible="false">
                                                    <ul>
                                                    <asp:Repeater ID="RP_Tab_LV3" runat="server" OnItemDataBound="RP_Tab_LV3_ItemDataBound" >
                                                        <ItemTemplate>
                                                            <li id="Tab_Li" runat="server">
                                                                <span ID="HID_CODEID" runat="server" visible="false" ><%#Container.DataItem("CODEID")%></span>
                                                                <asp:HyperLink ID="HPL_LV3" runat="server" Text='<%#Container.DataItem("TEXT")%>' NavigateUrl="#" />
                                                            </li>
                                                        </ItemTemplate>
                                                    </asp:Repeater>
                                                    </ul>
                                                </asp:PlaceHolder>
                                            </li>
                                        </ItemTemplate>
                                    </asp:Repeater>
                                </ul>
                            </asp:PlaceHolder>
                        </li>
                    </ItemTemplate>
                </asp:Repeater>
            </ul>
        </div>        
		<script src="<%=com.Azion.EloanUtility.UIUtility.getRootPath() + "/JS/prefixfree.js"%>" type="text/javascript"></script>
		<script src="<%=com.Azion.EloanUtility.UIUtility.getRootPath() + "/JS/jquery-1.7.1.min.js"%>" type="text/javascript"></script>
		<script>
		    $(document).ready(function () {
		        $("#accordian a").click(function () {
		            var link = $(this);
		            var closest_ul = link.closest("ul");
		            var parallel_active_links = closest_ul.find(".active")
		            var closest_li = link.closest("li");
		            var link_status = closest_li.hasClass("active");
		            var count = 0;

		            closest_ul.find("ul").slideUp(function () {
		                if (++count == closest_ul.find("ul").length)
		                    parallel_active_links.removeClass("active");
		            });

		            if (!link_status) {
		                closest_li.children("ul").slideDown();
		                closest_li.addClass("active");
		            }
		        })
		    })

		    var $sidebar = $("#accordian"),
                $window = $(window),
                offset = $sidebar.offset(),
                prevScrollTop = 0;

		    var fixmeTop = $sidebar.offset().top;
		    $(window).scroll(function () {
		        var currentScroll = $(window).scrollTop();

		        //alert(currentScroll);
		        //alert(fixmeTop);
		        //alert($sidebar.hasScrollBar());

		        if (currentScroll >= fixmeTop) {
		            $("#header-inner").css("display", "");
		            //$sidebar.css({
		            //    position: 'fixed',
		            //    top: '0',
		            //    marginLeft: '0px !important'
		            //});
		            //$sidebar.css('overflow-y', 'auto');
		            ////$sidebar.css('overflow', 'visible');
		            //$sidebar.css('max-height', '100%');
		            $sidebar.addClass("fixedTop");
		        } else {		            
		            if (currentScroll==0 && $sidebar.hasScrollBar()) {
		                $("#header-inner").css("display", "none");
		                //$sidebar.css({
		                //    position: 'fixed',
		                //    top: '0',
		                //    marginLeft: '0px !important'
		                //});
		                //$sidebar.css('overflow-y', 'auto');
		                ////$sidebar.css('overflow', 'visible');
		                //$sidebar.css('max-height', '100%');
		                $sidebar.addClass("fixedTop");
		            }
		            else {
		                $("#header-inner").css("display", "");
		                //$sidebar.css({
		                //    position: 'static'
		                //});
		                $sidebar.removeClass("fixedTop");
		            }
		        }

		    });

		    (function ($) {
		        $.fn.hasScrollBar = function () {
		            return this.get(0).scrollHeight > this.height();
		        }
		    })(jQuery);

        </script>
