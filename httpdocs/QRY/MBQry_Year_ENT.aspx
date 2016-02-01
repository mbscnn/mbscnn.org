<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="MBQry_Year_ENT.aspx.vb" Inherits="MBSC.MBQry_Year_ENT" %>

<!DOCTYPE html>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>MBSC年度活動一覽表</title>
    <!-- Bootstrap Core CSS -->
    <link href="<%=com.Azion.EloanUtility.UIUtility.getRootPathClient() + "/css/bootstrap.css"%>" rel="stylesheet" type="text/css" />
    <!-- Custom CSS -->
    <link href="<%=com.Azion.EloanUtility.UIUtility.getRootPathClient() + "/css/portfolio-item.css"%>" rel="stylesheet" type="text/css" />
    <link href="<%=com.Azion.EloanUtility.UIUtility.getRootPathClient() + "/css/font-awesome.min.css"%>" type="text/css" rel="stylesheet" />
    <link href="<%=com.Azion.EloanUtility.UIUtility.getRootPathClient() + "/css/animated.css"%>" rel="stylesheet" type="text/css" />
    <link href='http://fonts.googleapis.com/css?family=Open+Sans:300,400italic,400,600,700'   rel='stylesheet' type='text/css' />

    <!-- Script -->
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <Script lang="JavaScript" src="<%=com.Azion.EloanUtility.UIUtility.getRootPathClient() + "/js/JSUtil.js"%>" ></Script>
    <script lang="JavaScript" src="<%=com.Azion.EloanUtility.UIUtility.getRootPathClient() + "/js/KeyCheck.js"%>" ></script>
    <script lang="JavaScript" src="<%=com.Azion.EloanUtility.UIUtility.getRootPathClient() + "/js/json2.js"%>"></script>
    <script src="<%=com.Azion.EloanUtility.UIUtility.getRootPathClient() + "/js/jquery.js"%>" type="text/javascript"></script>
    <script src="<%=com.Azion.EloanUtility.UIUtility.getRootPathClient() + "/js/bootstrap.min.js"%>" type="text/javascript"></script>
    <script src="<%=com.Azion.EloanUtility.UIUtility.getRootPathClient() + "/js/scrolling-nav.js"%>" type="text/javascript"></script>
    <script src="<%=com.Azion.EloanUtility.UIUtility.getRootPathClient() + "/js/jquery.easing.min.js"%>" type="text/javascript"></script>
    <script type="text/javascript">
        jQuery(document).ready(function () {
            var offset = 220;
            var duration = 500;
            jQuery(window).scroll(function () {
                if (jQuery(this).scrollTop() > offset) {
                    jQuery('.back-to-top').fadeIn(duration);
                } else {
                    jQuery('.back-to-top').fadeOut(duration);
                }
            });

            jQuery('.back-to-top').click(function (event) {
                event.preventDefault();
                jQuery('html, body').animate({ scrollTop: 0 }, duration);
                return false;
            })
        });
        //jQuery(function ($) {

        //    $("table").removeAttr("style");
        //    $("table").addClass("table table-bordered");
        //});

    </script>

    <!--Responsive Table Style Start-->
    <style>
	    table { 
		    width: 100%; 
		    border-collapse: collapse; 
            /*background: -webkit-linear-gradient(top, rgba(255, 179, 51, 0.37) 0%, rgba(255, 255, 254, 0) 100%);
            background: linear-gradient(to bottom, rgb(249, 186, 128) 0%, rgba(249, 244, 241, 0.87) 100%);*/
            border-right: 1px solid #c3c3c3;
            border-bottom: 1px solid #c3c3c3;
	    }

	    tr:nth-of-type(odd) { 
            background: -webkit-linear-gradient(top, rgba(255, 255, 254, 0) 0%,rgba(255, 179, 51, 0.37)  100%);
            background: linear-gradient(to bottom, rgba(249, 244, 241, 0.87) 0%, rgb(249, 186, 128) 100%);
	    }

        tr
        {
            background: -webkit-linear-gradient(top, rgba(255, 179, 51, 0.37) 0%, rgba(255, 255, 254, 0) 100%);
            background: linear-gradient(to bottom, rgb(249, 186, 128) 0%, rgba(249, 244, 241, 0.87) 100%);
        }

	    th { 
		    background: #333; 
		    color: white; 
		    font-weight: bold; 
            font-size:14pt;
		    padding: 6px; 
		    border: 1px solid #ccc; 
		    text-align: center; 
	    }

	    td { 
		    padding: 6px; 
		    border: 1px solid #ccc; 
		    text-align: center; 
            font-size:12pt;
	    }
    </style>
    <!--[if !IE]><!-->
	<style>
	
	    /* 
	    Max width before this PARTICULAR table gets nasty
	    This query will take effect for any screen smaller than 760px
	    and also iPads specifically.
	    */
	    @media 
	    only screen and (max-width: 760px),
	    (min-device-width: 768px) and (max-device-width: 1024px)  {
	
		    /* Force table to not be like tables anymore */
		    table, thead, tbody, th, td, tr { 
			    display: block; 
		    }
		
		    /* Hide table headers (but not display: none;, for accessibility) */
		    thead tr { 
			    position: absolute;
			    top: -9999px;
			    left: -9999px;
		    }
		
		    tr { border: 1px solid #ccc; }
		
		    td { 
			    /* Behave  like a "row" */
			    border: none;
			    border-bottom: 1px solid #eee; 
			    position: relative;
			    padding-left: 50%; 
		    }
		
		    td:before { 
			    /* Now like a table header */
			    position: absolute;
			    /* Top/left values mimic padding */
			    top: 6px;
			    left: 6px;
			    width: 45%; 
			    padding-right: 10px; 
			    white-space: nowrap;
		    }
		
		    /*
		    Label the data
		    */
		    td:nth-of-type(1):before { content: "場序"; }
		    td:nth-of-type(2):before { content: "月份"; }
            td:nth-of-type(3):before { content: "活動名稱"; }
		    td:nth-of-type(4):before { content: "活動地點"; }
		    td:nth-of-type(5):before { content: "諮詢電話"; }
	    }
	
	    /* Smartphones (portrait and landscape) ----------- */
	    @media only screen
	    and (min-device-width : 320px)
	    and (max-device-width : 480px) {
		    body { 
			    padding: 0; 
			    margin: 0; 
			    width: 320px; }
		    }
	
	    /* iPads (portrait and landscape) ----------- */
	    @media only screen and (min-device-width: 768px) and (max-device-width: 1024px) {
		    body { 
			    width: 495px; 
		    }
	    }
	
	</style>
	<!--<![endif]-->
    <!--Responsive Table Style End-->
</head>
<body>
    <form id="form1" runat="server">
        <!-- #include virtual="~/inc/PageTab.inc" -->
        <!-- #include virtual="~/inc/Signin.inc" -->
        <div style="text-align:center">
            <h2 style="color:black; font-family: 標楷體;font-weight:bold">
                <asp:Literal ID="LTL_ENT_YEAR" runat="server" />
                MBSC年度活動一覽表
            </h2>
        </div>
        <!--錯誤訊息區-->
        <!-- #include virtual="~/inc/MBSCErrorMsg.inc" -->
            <table style="width:100%">
                <thead>
                    <tr>
                        <th >
                            場序
                        </th>
                        <th >
                            月份
                        </th>
                        <th >
                            活動名稱
                        </th>
                        <th >
                            活動地點
                        </th>
                        <th >
                            諮詢電話
                        </th>
                    </tr>
                </thead>
                <tr>
                    <th colspan="5" >
                        上半年度
                    </th>
                </tr>
                <asp:Repeater ID="RP_ENT_P6" runat="server">
                    <ItemTemplate>
                        <tr>
                            <td >
                                <!--場序-->
                                <asp:Literal ID="LTL_SEQ" runat="server" />
                            </td>
                            <td >
                                <!--月份-->
                                <asp:Literal ID="LTL_DATE" runat="server" />
                            </td>
                            <td >
                                <!--活動名稱-->
                                <asp:Literal ID="ENTNAME" runat="server" Text='<%#Container.DataItem("ENTNAME")%>' />
                            </td>
                            <td >
                                <!--活動地點-->
                                <asp:Literal ID="PLACE" runat="server" Text='<%#Container.DataItem("PLACE")%>' />
                            </td>
                            <td >
                                <!--諮詢電話-->
                                <asp:Literal ID="TEL" runat="server" Text='<%#Container.DataItem("TEL")%>' />
                            </td>
                        </tr>
                    </ItemTemplate>
                </asp:Repeater>
                <tr>
                    <th colspan="5">
                        下半年度
                    </th>
                </tr>
                <asp:Repeater ID="RP_ENT_E6" runat="server">
                    <ItemTemplate>
                        <tr>
                            <td >
                                <!--場序-->
                                <asp:Literal ID="LTL_SEQ" runat="server" />
                            </td>
                            <td >
                                <!--月份-->
                                <asp:Literal ID="LTL_DATE" runat="server" />
                            </td>
                            <td >
                                <!--活動名稱-->
                                <asp:Literal ID="ENTNAME" runat="server" Text='<%#Container.DataItem("ENTNAME")%>' />
                            </td>
                            <td >
                                <!--活動地點-->
                                <asp:Literal ID="PLACE" runat="server" Text='<%#Container.DataItem("PLACE")%>' />
                            </td>
                            <td >
                                <!--諮詢電話-->
                                <asp:Literal ID="TEL" runat="server" Text='<%#Container.DataItem("TEL")%>' />
                            </td>
                        </tr>
                    </ItemTemplate>
                </asp:Repeater>
                <tr>
                    <th colspan="5" style="text-align:left">
                        ※
                        以上活動，以學會官網公告為準。報名方式：一率採取網路報名。<BR/>
                        官網網址：<a href="http://mbscnn.org" >http://mbscnn.org</a>
                    </th>
                </tr>
                <tr>
                    <th colspan="5" style="text-align:left">
                        ※
                        各區教育中心所舉辦之課程活動，不在此列。請關注官網訊息公告。
                    </th>
                </tr>
            </table>
    </form>
</body>
</html>
