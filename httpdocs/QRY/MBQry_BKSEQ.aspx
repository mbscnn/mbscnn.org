<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="MBQry_BKSEQ.aspx.vb" Inherits="MBSC.MBQry_BKSEQ" %>

<!DOCTYPE html>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>護法會會員編號查詢</title>

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

    <style>
        .grid-container {
            /*width: 100%;
            max-width: 1200px;*/
            /*background-color: #FAEADD;*/
            background: -webkit-linear-gradient(top, rgba(255, 179, 51, 0.37) 0%, rgba(255, 255, 254, 0) 100%);
            background: linear-gradient(to bottom, rgb(249, 186, 128) 0%, rgba(249, 244, 241, 0.87) 100%);
            border-right: 1px solid #c3c3c3;
            border-bottom: 1px solid #c3c3c3;
        }

        .grid-container * {
          box-sizing: border-box;
        }

        /*所有div高度相同*/
        /*.row{
            overflow: hidden; 
        }*/

        .row:before,
        .row:after {
          content: "";
          display: table;
          clear: both;
        }
         
        .thm
        {
            /*background-color: #003040;
            background: -webkit-linear-gradient(top, rgba(255, 255, 254, 0) 100% ,rgba(255, 179, 51, 0.37) 0%);
            background: linear-gradient(to bottom, rgba(249, 244, 241, 0.87) 100%, rgb(249, 186, 128) 0%);*/
            /*background:rgba(249, 186, 128, 0.5);*/
            background:transparent;
            font-weight:bold;
            padding: 7px;
        }
		
        @media all and (max-width:767px) {
            .row [class*="col-"] {
                font-size: 11pt;
                /*float: left;*/
                /*min-height: 1px;*/
                /*width: 16.66%;*/
                padding: 3px;
                /*background: #eee;*/
                background:transparent;
                border-top: 1px solid #c3c3c3;
                /*border-left: 1px solid #c3c3c3;*/
                text-align:left;

                /*所有div高度相同*/
                /*margin-bottom: -99999px;
                padding-bottom: 99999px;*/
            }

            /*.form-control
            {
                margin-bottom:3px;
            }*/
        }

        /* Small devices (tablets, 768px and up) */
        @media all and (min-width: 768px) {
            .row [class*="col-"] {
                font-size: 11pt;
                /*float: left;*/
                /*min-height: 1px;*/
                /*width: 16.66%;*/
                padding: 5px;
                /*background: #eee;*/
                background:transparent;
                border-top: 1px solid #c3c3c3;
                /*border-left: 1px solid #c3c3c3;*/
                text-align:left;

                /*所有div高度相同*/
                /*margin-bottom: -99999px;
                padding-bottom: 99999px;*/
            }

            /*.form-control
            {
                margin-bottom:5px;
            }*/
        }

        /* Medium devices (desktops, 992px and up) */
        @media all and (min-width: 992px) {
            .row [class*="col-"] {
                font-size: 12pt;
                /*float: left;*/
                /*min-height: 1px;*/
                /*width: 16.66%;*/
                padding: 7px;
                /*background: #eee;*/
                background:transparent;
                border-top: 1px solid #c3c3c3;
                /*border-left: 1px solid #c3c3c3;*/
                text-align:left;

                /*所有div高度相同*/
                /*margin-bottom: -99999px;
                padding-bottom: 99999px;*/
            }

            /*.form-control
            {
                margin-bottom:7px;
            }*/
        }

        /* Large devices (large desktops, 1200px and up) */
        @media all and (min-width: 1200px) {
            .row [class*="col-"] {
                font-size: 14pt;
                /*float: left;*/
                /*min-height: 1px;*/
                /*width: 16.66%;*/
                padding: 9px;
                /*background: #eee;*/
                background:transparent;
                border-top: 1px solid #c3c3c3;
                /*border-left: 1px solid #c3c3c3;*/
                text-align:left;

                /*所有div高度相同*/
                /*margin-bottom: -99999px;
                padding-bottom: 99999px;*/
            }

            /*.form-control
            {
                margin-bottom:9px;
            }*/
        }
    </style>

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
		    padding: 6px; 
		    border: 1px solid #ccc; 
		    text-align: center; 
	    }

	    td { 
		    padding: 6px; 
		    border: 1px solid #ccc; 
		    text-align: center; 
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
		    td:nth-of-type(1):before { content: "姓名"; }
		    td:nth-of-type(2):before { content: "通訊地址"; }
            td:nth-of-type(3):before { content: "e-mail"; }
		    td:nth-of-type(4):before { content: "手機/電話"; }
		    td:nth-of-type(5):before { content: "護法會員編號"; }
		    td:nth-of-type(6):before { content: "是否為會員"; }
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
                社團法人台灣佛陀原始正法學會
            </h2>
            <h3 style="color:black; font-family: 標楷體;font-weight:bold">
                護法會員編號查詢
            </h3>
        </div>
        <!--錯誤訊息區-->
        <!-- #include virtual="~/inc/MBSCErrorMsg.inc" -->
        <div id="DIV_Paras" runat="server" class="container-fluid grid-container" >
            <div class="row">
                <div class="col-xs-4 col-md-1 thm">
                    <asp:RadioButton ID="rbt_Cel" runat="server" Text="手機：" GroupName="SelectType" />
                </div>
                <div class="col-xs-8 col-md-3">
                    <asp:TextBox ID="TXT_MB_MOBIL" runat="server" class="form-control" />
                </div>
                <div class="col-xs-4 col-md-1 thm">
                    <asp:RadioButton ID="rbt_Tel" runat="server" Text="電話：" GroupName="SelectType" />
                </div>
                <div class="form-group col-xs-8 col-md-3">
                        <asp:TextBox ID="TXT_MB_TEL_ZIP" runat="server" class="form-control" Style="display:inline;width:20%" MaxLength="2" />
                        ─
                        <asp:TextBox ID="TXT_MB_TEL" runat="server" class="form-control" Style="display:inline;width:60%" MaxLength="10" />
                </div>
                <div class="col-xs-4 col-md-1 thm">
                    <asp:RadioButton ID="rbt_Name" runat="server" Text="姓名：" GroupName="SelectType" onclick="alert('需輸入全名');" />
                </div>
                <div class="col-xs-8 col-md-3">
                    <asp:TextBox ID="TXT_MB_NAME" runat="server" class="form-control" />
                </div>
            </div>
            <div class="row">
                <div class="col-xs-12 col-md-12" style="text-align:center">
                    <asp:Button ID="btn_Qry" runat="server" Text="確定" CssClass="btn btn-info" />
                </div>
            </div>
            <div class="row">
                <div class="col-xs-12 col-md-12">
                    16碼護法會員編號用途
                    <ol type="1">                
                        <li>台銀代扣善款授權書</li>
                        <li>上台銀網銀網路ATM轉帳給佛陀正法中心善款(不須付轉帳費用)</li>
                        <li>到任何ATM轉帳給佛陀正法中心善款(捐贈者須付轉帳費用15元)</li>
                        <li>報名佛陀正法中心各類課程</li>
                        <li>提供捐贈善款及上課報名查詢</li>
                    </ol>
                </div>
            </div>
        </div>
        <asp:PlaceHolder ID="PLH_DATA" runat="server" Visible="false">
        <table>
            <thead>
                <tr>
                    <th>姓名</th>
                    <th>通訊地址</th>
                    <th>e-mail</th>
                    <th>手機/電話</th>
                    <th>護法會員編號</th>
                    <th>是否為會員</th>
                </tr>
            </thead>
            <tbody>
                <asp:Repeater ID="RP_BKSEQ" runat="server" >
                    <ItemTemplate>
                        <tr>
                            <td>
                                <!--姓名-->
                                <asp:Literal ID="LTL_MB_NAME" runat="server" Text='<%#Container.DataItem("MB_NAME")%>' />
                                &nbsp;
                            </td>
                            <td>
                                <!--通訊地址-->
                                <asp:Literal ID="MB_ADDR" runat="server"  />
                                &nbsp;
                            </td>
                            <td>
                                <!--e-mail-->
                                <asp:Literal ID="MB_EMAIL" runat="server" Text='<%#Container.DataItem("MB_EMAIL")%>' />
                            </td>
                            <td>
                                <!--手機/電話-->
                                <asp:Literal ID="LTL_MB_MOBIL" runat="server" />
                                &nbsp;
                            </td>
                            <td>
                                <!--護法會員編號-->
                                <asp:Literal ID="LTL_MB_BKSEQ" runat="server" Text='<%#Container.DataItem("MB_BKSEQ")%>' />
                                &nbsp;
                            </td>
                            <td>
                                <!--是否為會員-->
                                <asp:Literal ID="LTL_MB_ACCT" runat="server" />
                                <asp:Button ID="btnSign" runat="server" Text="成為新會員" CssClass="bt btn-info" Visible="false" CommandName="SIGN" CommandArgument='<%#Container.DataItem("MB_MEMSEQ")%>' />
                                &nbsp;
                            </td>
                        </tr>
                    </ItemTemplate>
                </asp:Repeater>
            </tbody>
        </table>
        </asp:PlaceHolder>
        <div id="DIV_NODATA" runat="server" visible="false" class="container-fluid grid-container" >
            <div class="row">
                <div class="col-xs-12 col-md-12">
                   <span style="color:red;font-weight:bold">查無資料,請</span>
                    <ol type="1">
                        <li>在佛陀原始正法中心官網上方<asp:Button ID="btnSingin" runat="server" Text="會員登入" CssClass="btn btn-info" />成為新會員</li>
                        <li>回e-mail信箱確認會員信函</li>
                        <li>會員信函確認完成,到佛陀原始正法中心官網<span style="color:blue;font-weight:bold">點選會員系統</span>填入<span style="color:blue;font-weight:bold">入會申請單</span>取得護法會員編號</li>
                        <li>若不清楚請來電中心02-23627968查問</li>
                    </ol>
                </div>
            </div>
        </div>
        <div id="DIV_BACK" runat="server" visible="false" style="text-align:center" >
            <asp:Button ID="btnBack" runat="server" Text="重新輸入查詢條件" CssClass="btn btn-info" />
        </div>

    </form>
</body>
</html>
