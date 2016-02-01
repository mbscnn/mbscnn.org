<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="MBQry_BKFC.aspx.vb" Inherits="MBSC.MBQry_BKFC" %>

<!DOCTYPE html>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>護法會員捐贈功德暨學習護照</title>
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
    <script type="text/javascript" language="JavaScript">
        //DataGrid Radiobutton防止多選
        function RadioButtonSelect(aspRadioButtonID) {
            re = new RegExp('RB_CHOOSE')
            for (i = 0; i < document.forms[0].elements.length; i++) {

                elm = document.forms[0].elements[i]

                if (elm.type == 'radio') {
                    if (re.test(elm.name)) {
                        elm.checked = false
                    }
                }
            }
            aspRadioButtonID.checked = true;
        };
    </script>
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
                護法會員捐贈功德暨學習護照
            </h3>
        </div>
        <!--錯誤訊息區-->
        <!-- #include virtual="~/inc/MBSCErrorMsg.inc" -->
        <asp:PlaceHolder ID="PLH_MEMBER" runat="server" Visible="false" >
            <div class="table-responsive" >
                <table class="table CRTable_Top" width="100%" cellspacing="0">
                    <tr>
                        <td width="10%" class="th1c_b">
                            選取
                        </td>
                        <td width="15%" class="th1c_b">
                            姓名
                        </td>
                        <td width="30%" class="th1c_b">
                            通訊地址
                        </td>
                        <td width="15%" class="th1c_b">
                            e-mail
                        </td>
                        <td width="15%" class="th1c_b">
                            手機/電話
                        </td>
                        <td width="15%" class="th1c_b">
                            護法會員編號
                        </td>
                    </tr>
                    <asp:Repeater ID="RP_MEMBER" runat="server" >
                        <ItemTemplate>
                            <tr>
                                <td width="10%" class="td2c_b">
                                    <!--選取-->
                                    <asp:RadioButton ID="RB_CHOOSE" runat="server" onclick="RadioButtonSelect(this);"  />
                                </td>
                                <td width="15%" class="td2c_b">
                                    <!--姓名-->
                                    <asp:Literal ID="LTL_MB_NAME" runat="server" Text='<%#Container.DataItem("MB_NAME")%>' />
                                </td>
                                <td width="30%" class="td2_b">
                                    <!--通訊地址-->
                                    <asp:Literal ID="LTL_MB_ADDR" runat="server" />
                                </td>
                                <td width="15%" class="td2_b">
                                    <!--e-mail-->
                                    <asp:Literal ID="LTL_MB_EMAIL" runat="server" Text='<%#Container.DataItem("MB_EMAIL")%>' />
                                </td>
                                <td width="15%" class="td2_b">
                                    <!--手機/電話-->
                                    <asp:Literal ID="LTL_MB_TEL" runat="server" />
                                </td>
                                <td width="15%" class="td2c_b">
                                    <!--護法會員編號-->
                                    <asp:Literal ID="LTL_MB_BKSEQ" runat="server" Text='<%#Container.DataItem("MB_BKSEQ")%>' />
                                </td>
                            </tr>
                        </ItemTemplate>
                    </asp:Repeater>
                </table>
            </div>
            <div style="text-align:center">
                <asp:Button ID="btn_Sign_Qry" runat="server" Text="確定" CssClass="btn btn-info" />
            </div>
        </asp:PlaceHolder>
        <asp:PlaceHolder ID="PLH_Paras" runat="server" >
            <div class="container-fluid grid-container">
                <div class="row">
                    <div class="col-md-12 thm">
                        請先輸入
                    </div>
                </div>
                <div class="row">
                    <div class="col-md-2 thm">
                        護法會員編號
                    </div>
                    <div class="col-md-10">
                        <asp:TextBox ID="MB_BKSEQ" runat="server" class="form-control" />
                    </div>
                </div>
                <div class="row">
                    <div class="col-md-2 thm">
                        查詢類別
                    </div>
                    <div class="col-md-10">
                        <asp:RadioButtonList ID="RBL_QRYTYP" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                            <asp:ListItem Text="捐贈" Value="1"></asp:ListItem>
                            <asp:ListItem Text="課程" Value="2"></asp:ListItem>
                        </asp:RadioButtonList>
                    </div>
                </div>
            </div>
            <div style="text-align:center">
                <asp:Button ID="btn_Qry" runat="server" Text="查詢" CssClass="btn btn-info" />
            </div>
        </asp:PlaceHolder>

            <asp:PlaceHolder ID="PLH_MB_MEMREV" runat="server" Visible="false">
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
		                td:nth-of-type(1):before { content: "繳款日期"; }
		                td:nth-of-type(2):before { content: "繳款金額"; }
                        td:nth-of-type(3):before { content: "功德項目"; }
		                td:nth-of-type(4):before { content: "會員類別"; }
		                td:nth-of-type(5):before { content: "捐贈管道"; }
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

                <table>
                    <thead>
                        <tr>
                            <th>繳款日期</th>
                            <th>繳款金額</th>
                            <th>功德項目</th>
                            <th>會員類別</th>
                            <th>捐贈管道</th>
                        </tr>
                    </thead>
                    <tbody>
                        <asp:Repeater ID="RP_MB_MEMREV" runat="server">
                            <ItemTemplate>
                                <tr>
                                    <td>
                                        <!--繳款日期-->
                                        <asp:Literal ID="MB_TX_DATE" runat="server" />&nbsp;
                                    </td>
                                    <td>
                                        <!--繳款金額-->
                                        <asp:Literal ID="MB_TOTFEE" runat="server" />&nbsp;
                                    </td>
                                    <td>
                                        <!--功德項目-->
                                        <asp:Literal ID="MB_ITEMID" runat="server" />&nbsp;
                                    </td>
                                    <td>
                                        <!--會員類別-->
                                        <asp:Literal ID="MB_MEMTYP" runat="server" />&nbsp;
                                    </td>
                                    <td>
                                        <!--捐贈管道-->
                                        &nbsp;
                                    </td>
                                </tr>
                            </ItemTemplate>
                        </asp:Repeater>
                    </tbody>
                </table>
                <div style="text-align:center">
                    <asp:Button ID="btnReQry_1" runat="server" Text="取消重新查詢" CssClass="btn btn-info" />
                </div>
            </asp:PlaceHolder>

            <asp:PlaceHolder ID="PLH_MB_MEMCLASS" runat="server" Visible="false">
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
		                td:nth-of-type(1):before { content: "課程編號"; }
		                td:nth-of-type(2):before { content: "梯次"; }
                        td:nth-of-type(3):before { content: "地點"; }
		                td:nth-of-type(4):before { content: "課程起訖日"; }
		                td:nth-of-type(5):before { content: "課程名稱"; }
		                td:nth-of-type(6):before { content: "指導老師"; }
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

                <table>
                    <thead>
                        <tr>
                            <th>課程編號</th>
                            <th>梯次</th>
                            <th>地點</th>
                            <th>課程起訖日</th>
                            <th>課程名稱</th>
                            <th>指導老師</th>
                        </tr>
                    </thead>
                    <tbody>
                        <asp:Repeater ID="RP_MB_MEMCLASS" runat="server" >
                            <ItemTemplate>
                                <tr>
                                    <td>
                                        <!--課程編號-->
                                        <asp:Literal ID="MB_SEQ" runat="server" Text='<%#Container.DataItem("MB_SEQ")%>' />&nbsp;
                                    </td>
                                    <td>
                                        <!--梯次-->
                                        <asp:Literal ID="LTL_MB_BATCH" runat="server" />&nbsp;
                                    </td>
                                    <td>
                                        <!--地點-->
                                        <asp:Literal ID="MB_PLACE" runat="server" />&nbsp;
                                    </td>
                                    <td>
                                        <!--課程起訖日-->
                                        <asp:Literal ID="MB_SDATE" runat="server" />&nbsp;
                                    </td>
                                    <td>
                                        <!--課程名稱-->
                                        <asp:Literal ID="MB_CLASS_NAME" runat="server" />&nbsp;
                                    </td>
                                    <td>
                                        <!--指導老師-->
                                        <asp:Literal ID="MB_TEACHER" runat="server" />&nbsp;
                                    </td>
                                </tr>
                            </ItemTemplate>
                        </asp:Repeater>
                    </tbody>
                </table>
                <div style="text-align:center">
                    <asp:Button ID="btnReQry_2" runat="server" Text="取消重新查詢" CssClass="btn btn-info" />
                </div>
            </asp:PlaceHolder>
    </form>
</body>
</html>
