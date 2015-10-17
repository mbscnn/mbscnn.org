<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="MBMnt_Sign_01_v01.aspx.vb" Inherits="MBSC.MBMnt_Sign_01_v01" %>

<!DOCTYPE html>

<html>
<head runat="server">
    <title></title>
    <script type="text/javascript" language="JavaScript" src="<%=com.Azion.EloanUtility.UIUtility.getRootPathClient() + "/js/jquery-1.7.1.min.js"%>"></script>
    <script type="text/javascript" language="JavaScript">
        //DataGrid Radiobutton防止多選
        function RadioButtonSelect(aspRadioButtonID) {
            re = new RegExp('rbData')
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
        //出家眾
        function CheckMonk(thisobj) {
            if (thisobj == document.getElementById("rbt_MB_MONK_N")) {
                document.getElementById('tr_MB_MONK_1').style.display = 'none';
                document.getElementById('tr_MB_MONK_2').style.display = 'none';
                document.getElementById('tr_MB_MONK_3').style.display = 'none';
                document.getElementById('tr_MB_MONK_4').style.display = 'none';
                return true;
            } else {
                document.getElementById('tr_MB_MONK_1').style.display = '';
                document.getElementById('tr_MB_MONK_2').style.display = '';
                document.getElementById('tr_MB_MONK_3').style.display = '';
                document.getElementById('tr_MB_MONK_4').style.display = '';
                return true;
            }

            return false;
        };
    </script>
    <!-- #include virtual="~/inc/MBSCCSS.inc" -->
    <!--CSS-->
    <!-- #include virtual="~/inc/MBSCJS.inc" -->
    <!--JS-->
    <style>
        /*div.row{
          border: 1px solid;
          border-bottom: 0px;
        }

        .container div.row:last-child {
          border-bottom: 1px solid;
        }*/

        /*.my-container [class^="col-"] {
            padding-top: 10px;
            padding-bottom: 10px;
            background-color: #eee;
            border: 1px solid #c3c3c3;
            background-color: rgba(86,61,124,.15);
            border: 1px solid rgba(86,61,124,.2);
        }*/

        /*.row {
            border: 1px solid #c2c2c2;
        }*/

        /*.row [class^="col-"] {
            border-right: 1px solid #c2c2c2;
        }

        .row:last-child {
            border-right: none;
        }*/

        /*.row + .row {
            border-top: 0;
        }*/

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
</head>
<body>
    <form id="form1" runat="server">
        <!-- #include virtual="~/inc/PageTab.inc" -->
        <!-- #include virtual="~/inc/Signin.inc" -->
        <div style="text-align:center">
            <h2 style="color: #800080; font-family: 標楷體">
                MBSC報名表
            </h2>
        </div>
        <!--錯誤訊息區-->
        <!-- #include virtual="~/inc/MBSCErrorMsg.inc" -->
        <div id="tb_Page1" runat="server" class="container-fluid grid-container">
            <div class="row">
                <div class="col-md-12 thm">
                    帳號：<asp:Label ID="lbl_USERID" runat="server" Text="" ForeColor="Blue"></asp:Label>
                    <br />
                    <span style="color: #FF0000">(若非報名本人請輸入報名者手機或電話擇一，將選出報名者資料)</span>
                </div>
            </div>
            <div class="row">
                <div class="col-md-2 thm">
                    <asp:RadioButton ID="rbt_Cel" runat="server" Text="手機：" GroupName="SelectType" />
                </div>
                <div class="col-md-4">
                    <asp:TextBox ID="txt_Cel" runat="server" class="form-control" />
                </div>
                <div class="col-md-2 thm">
                    <asp:RadioButton ID="rbt_Tel" runat="server" Text="電話：" GroupName="SelectType" />
                </div>
                <div class="col-md-4">
                    <asp:TextBox ID="txt_Tel_Zip" runat="server" Columns="2" MaxLength="2"></asp:TextBox>
                    &nbsp;─&nbsp;
                    <asp:TextBox ID="txt_Tel" runat="server" Columns="10" MaxLength="10"></asp:TextBox>
                </div>
            </div>
        </div>
        <div id="tb_Page1_btn" runat="server" class="text-center">
            <asp:Button ID="btn_Qry" runat="server" Text="確定" CssClass="btn btn-info" />
        </div>

        <asp:PlaceHolder ID="PLH_MAIL_SAME" runat="server" Visible="false">
            <div class="container-fluid grid-container">
            	<div class="row">
            		<div class="col-xs-1 col-md-1 thm">
            			點選
            		</div>
            		<div class="col-xs-3 col-md-3 thm">
            			姓名
            		</div>
            		<div class="col-xs-6 col-md-6 thm">
            			通訊地址
            		</div>
            		<div class="col-xs-2 col-md-2 thm">
            			會員編號
            		</div>
            	</div>
            	<asp:Repeater ID="RP_MAIL_SAME" runat="server">
            		<ItemTemplate>
            			<div class="row">
		            		<div class="col-xs-1 col-md-1">
		            			<!--點選-->
                                <asp:RadioButton ID="RB_CHOOSE" runat="server" AutoPostBack="true" OnCheckedChanged="RB_CHOOSE_OnCheckedChanged" />
                                <input type="hidden" id="HID_MB_MEMSEQ" runat="server" value='<%#Container.DataItem("MB_MEMSEQ")%>' />		            					
		            		</div>
		            		<div class="col-xs-3 col-md-3">
		            			<!--姓名--> 
		            			<%#Container.DataItem("MB_NAME")%>
		            		</div>
		            		<div class="col-xs-6 col-md-6">
		            			<!--通訊地址-->
		            			<%# Container.DataItem("MB_CITY")%><%# Container.DataItem("MB_VLG")%><%#Container.DataItem("MB_ADDR")%>
		            		</div>
		            		<div class="col-xs-2 col-md-2">
		            			<!--會員編號-->
		            			<%#getMB_MEMSEQ(Container.DataItem("MB_MEMSEQ"), Container.DataItem("MB_AREA"))%>
		            		</div>            						
            			</div>
            		</ItemTemplate>
            	</asp:Repeater>
            </div>
            <div class="text-center">
                <asp:Button ID="btnModify" runat="server" CssClass="btn btn-info" Text="報名" />
            </div>
        </asp:PlaceHolder>
        <div id="dgRpt_Page2" runat="server" visible="false" class="table-responsive grid-container">
            <table class="table">
                <tbody>
                    <tr  >
                        <td class="th1c_b" style="background:transparent">
                            點選
                        </td>
                        <td class="th1c_b" style="background:transparent">
                            法名/姓名
                        </td>
                        <td class="th1c_b" style="background:transparent">
                            通訊地址
                        </td>
                        <td class="th1c_b" style="background:transparent">
                            會員編號
                        </td>
                    </tr>				
                    <asp:Repeater ID="RP_Page2" runat="server">
                        <ItemTemplate>
                            <tr>
                                <td class="td2c_b" style="background:transparent">
                                    <!--點選-->
                                    <asp:RadioButton ID="rbData" runat="server" onclick="RadioButtonSelect(this);" />
                                </td>
                                <td class="td2c_b" style="background:transparent">
                                    <!--法名/ 姓名-->
                                    <asp:Label ID="lbl_MB_NAME" runat="server" Text='<%#Container.DataItem("MB_NAME")%>'></asp:Label>
                                </td>
                                <td class="td2_b" style="background:transparent">
                                    <!--通訊地址-->
                                    <asp:Label ID="lbl_MB_CITY" runat="server" Text='<%#Container.DataItem("MB_CITY")%>'></asp:Label>
                                    <asp:Label ID="lbl_MB_VLG" runat="server" Text='<%#Container.DataItem("MB_VLG")%>'></asp:Label>
                                    <asp:Label ID="lbl_MB_ADDR" runat="server" Text='<%#Container.DataItem("MB_ADDR")%>'></asp:Label>
                                </td>
                                <td class="td2c_b" style="background:transparent">
                                    <!--會員編號-->
                                    <asp:Label ID="lbl_MB_AREA" runat="server" Text='<%#Container.DataItem("MB_AREA")%>'></asp:Label>
                                    -
                                    <asp:Label ID="lbl_MB_MEMSEQ" runat="server" Text='<%#Container.DataItem("MB_MEMSEQ")%>'></asp:Label>
                                </td>
                            </tr>
                        </ItemTemplate>
                    </asp:Repeater>
                </tbody>
            </table>
        </div>
        <div id="tb_Page2_Button" runat="server" class="text-center" visible="false" >
            <asp:Button ID="btn_Confirm" runat="server" Text="確定" CssClass="btn btn-info" />
        </div>
						
		<div id="tb_Page3" runat="server" class="container-fluid grid-container" visible="false" >
			<div class="row">
				<div class="col-md-12 thm"  >
                    會員編號 : 
                    <asp:Label ID="lbl_MB_AREA" runat="server" Text="" />-
                    <asp:Label ID="lbl_MB_MEMSEQ" runat="server" Text="" />
                    &nbsp;基本資料									
				</div>
			</div>
			<div class="row">
				<div class="col-md-2 thm">
					課程名稱
				</div>
				<div class="col-md-10">
					<asp:Label ID="lbl_Class" runat="server" Text="" />
				</div>
			</div>
			<div class="row">
				<div class="col-md-2 thm">
					<span style="color:Red">*</span>姓名
				</div>
				<div class="col-md-4">
					<asp:TextBox ID="txt_MB_NAME" runat="server" class="form-control"  />
				</div>
				<div class="col-md-2 thm">
					<span style="color:Red">*</span>
					性別
				</div>
				<div class="col-md-4">
                    <asp:RadioButtonList ID="rbtList_MB_SEX" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow"   >
                        <asp:ListItem Value="1" Text="男" class="radio-inline"  ></asp:ListItem>
                        <asp:ListItem Value="2" Text="女" class="radio-inline" ></asp:ListItem>
                    </asp:RadioButtonList>									
				</div>
			</div>
			<div class="row">
				<div class="col-md-2 thm">
					<span style="color:Red">*</span>
					出生年月日
				</div>
				<div class="col-md-4 text-right" id="TD_Y_1" runat="server" >
                    西元
                    <asp:TextBox ID="txt_MB_BIRTH_YYY" runat="server" Columns="4" ></asp:TextBox>年&nbsp;
                    <asp:TextBox ID="txt_MB_BIRTH_MM" runat="server" Columns="2" ></asp:TextBox>月&nbsp;
                    <asp:TextBox ID="txt_MB_BIRTH_DD" runat="server" Columns="2" ></asp:TextBox>日&nbsp;									
				</div>            
				<div class="col-md-2 thm" id="TD_G_1_1" runat="server" >
					<span style="color:Red">*</span>
					出家眾
				</div>
				<div class="col-md-4" id="TD_G_1_2" runat="server" >
                    <asp:RadioButton ID="rbt_MB_MONK_Y" runat="server" Text="是" GroupName="MB_MONK" onclick="CheckMonk(this);"></asp:RadioButton>
                    <asp:RadioButton ID="rbt_MB_MONK_N" runat="server" Text="否" GroupName="MB_MONK" onclick="CheckMonk(this);"></asp:RadioButton>									
				</div>
			</div>
			<div id="tr_MB_MONK_1" runat="server" class="row" style="display:none" >
				<div class="col-md-2 thm">
					法名
				</div>
				<div class="col-md-4">
					<asp:TextBox ID="txt_MB_MONKNAME" runat="server" class="form-control" />
				</div>
				<div class="col-md-2 thm">
					剃度/皈依恩師/戒師
				</div>
				<div class="col-md-4">
					<asp:TextBox ID="txt_MB_MONKTECH" runat="server" class="form-control" />
				</div>
			</div>
			<div id="tr_MB_MONK_2" runat="server" class="row" style="display:none">
				<div class="col-md-2 thm">
					傳承
				</div>
				<div class="col-md-4">
					<asp:DropDownList ID="dd_MB_EDUTYPE" runat="server" CssClass="CtlFnt" />
				</div>
				<div class="col-md-2 thm">
					常住/親近道場
				</div>
				<div class="col-md-4">
					<asp:TextBox ID="txt_MB_MONKPLACE" runat="server" class="form-control" />
				</div>
			</div>
			<div id="tr_MB_MONK_3" runat="server" class="row" style="display:none">
				<div class="col-md-2 thm">
					戒別
				</div>
				<div class="col-md-4">
					<asp:DropDownList ID="dd_MB_MONKTYPE" runat="server" CssClass="CtlFnt" />
				</div>
				<div class="col-md-2 thm">
					受戒日期
				</div>
				<div class="col-md-4">
					西元
					<asp:TextBox ID="txt_MB_MONKDATE_YYY" runat="server" Columns="4" ></asp:TextBox>年&nbsp;
					<asp:TextBox ID="txt_MB_MONKDATE_MM" runat="server" Columns="2" ></asp:TextBox>月&nbsp;
					<asp:TextBox ID="txt_MB_MONKDATE_DD" runat="server" Columns="2" ></asp:TextBox>日
				</div>
			</div>
			<div id="tr_MB_MONK_4" runat="server" class="row" style="display:none">
				<div class="col-md-2 thm">
					受戒地點
				</div>
				<div class="col-md-10">
					<asp:TextBox ID="txt_MB_MONKPLACE1" runat="server" class="form-control" />
				</div>
			</div>
			<div class="row" >
				<div class="col-md-2 thm">
					<span style="color:Red">*</span>
					手機
				</div>
				<div class="col-md-4">
					<asp:TextBox ID="txt_MB_MOBIL" runat="server" class="form-control" />
				</div>
				<div class="col-md-2 thm">
					<span style="color:Red">*</span>
					電話
				</div>
				<div class="col-md-4">
					<asp:TextBox ID="txt_MB_TEL_ZIP" runat="server" Columns="2"  MaxLength="2" />
					&nbsp;─&nbsp;
					<asp:TextBox ID="txt_MB_TEL" runat="server" Columns="10" />
				</div>
			</div>
			<div class="row" >
				<div class="col-md-2 thm">
					<span style="color:Red">*</span>
					E-mail
				</div>
				<div id="TD_Y_2" runat="server" class="col-md-4">
					<asp:TextBox ID="txt_MB_EMAIL" runat="server" class="form-control" />
				</div>
				<div id="TD_G_2_1" runat="server"  class="col-md-2 thm">
					<span style="color:Red">**</span>
					身分證字號
				</div>
				<div id="TD_G_2_2" runat="server"  class="col-md-4">
					<asp:TextBox ID="txt_MB_ID" runat="server" class="form-control" />
				</div>
			</div>
			<div class="row" >
				<div class="col-md-2 thm">
					<span style="color:Red">*</span>
					參加動機或目的
				</div>
				<div class="col-md-10">
					<asp:TextBox ID="MB_OBJECT" runat="server" class="form-control" />
				</div>
			</div>
			<div class="row" >
				<div class="col-md-2 thm">
					<span style="color:Red">*</span>
					是否曾參加過本中心課程?
				</div>
				<div class="col-md-10">
                    <asp:RadioButtonList ID="JOINMBSC" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow" >
                        <asp:ListItem Text="是" Value="Y" />
                        <asp:ListItem Text="否" Value="N" />
                    </asp:RadioButtonList>									
				</div>
			</div>
			<div class="row" >
				<div class="col-md-2 thm">
					學歷
				</div>
				<div class="col-md-4">
					<asp:DropDownList ID="dd_MB_EDU" runat="server" CssClass="CtlFnt" />
				</div>
				<div class="col-md-2 thm">
					介紹人/訊息來源
				</div>
				<div class="col-md-4">
					<asp:TextBox ID="txt_MB_REFER" runat="server" class="form-control" />
				</div>
			</div>
			<div id="TR_G_3" runat="server" class="row" >
				<div class="col-md-2 thm">
					<span style="color:Red">*</span>
					通訊地址
				</div>
				<div class="col-md-10">
                    <asp:DropDownList ID="dd_MB_CITY" runat="server" AutoPostBack="true" CssClass="CtlFnt" />
                    &nbsp;
                    <asp:DropDownList ID="dd_MB_VLG" runat="server" CssClass="CtlFnt" /> 
                    &nbsp;
                    <asp:TextBox ID="txt_MB_ADDR" runat="server" class="form-control" />					
				</div>
			</div>
			<div id="TR_G_4" runat="server" class="row" >
				<div class="col-md-2 thm">
					戶籍地址
					<asp:CheckBox ID="cb_Ditto" runat="server" Text="同上" AutoPostBack="True" />
				</div>
				<div class="col-md-10">
                    <asp:DropDownList ID="dd_MB_CITY1" runat="server" AutoPostBack="true" CssClass="CtlFnt" />
                    &nbsp;
                    <asp:DropDownList ID="dd_MB_VLG1" runat="server" CssClass="CtlFnt" />
                    &nbsp;
                    <asp:TextBox ID="txt_MB_ADDR1" runat="server" class="form-control" />						
				</div>
			</div>
			<div id="TR_G_5" runat="server" class="row" >
				<div class="col-md-2 thm">
					語言
				</div>
				<div class="col-md-10">
                    <asp:Repeater ID="dl_MB_LANG" runat="server">
                        <ItemTemplate>
                            <asp:HiddenField ID="hid_MB_LANG" runat="server" Value='<%#Container.DataItem("VALUE")%>' />
                            <asp:CheckBox ID="cb_MB_LANG" runat="server" Text='<%#Container.DataItem("TEXT")%>' CssClass="CtlFnt" >
                            </asp:CheckBox>
                            <asp:TextBox ID="txt_MB_LANG" runat="server" Visible="false" class="form-control" ></asp:TextBox>
                        </ItemTemplate>
                    </asp:Repeater>
				</div>
			</div>
			<div id="TR_G_6" runat="server" class="row" >
				<div class="col-md-2 thm">
					專長
				</div>
				<div class="col-md-4">
                    <asp:DropDownList ID="dd_MB_SPECIAL" runat="server" CssClass="CtlFnt" />
				</div>
				<div class="col-md-2 thm">
					職業
				</div>
				<div class="col-md-4">
                    <asp:DropDownList ID="dd_MB_JOB" runat="server" CssClass="CtlFnt" />
                    <BR/>
                    職稱
                    <BR/>
                    <asp:DropDownList ID="dd_MB_JOBTITLE" runat="server" CssClass="CtlFnt" />							
				</div>
			</div>
			<div class="row" >
				<div class="col-md-2 thm">
					<span style="color:Red">*</span>
					宗教信仰
				</div>
				<div id="TD_Y_7" runat="server"  class="col-md-4">
					<asp:DropDownList ID="dd_MB_RELIGION" runat="server" CssClass="CtlFnt" />
				</div>
				<div id="TD_G_7_1" runat="server" class="col-md-2 thm">
					<span style="color:Red">**</span>
					打鼾
				</div>
				<div id="TD_G_7_2" runat="server"  class="col-md-4">
                    <asp:RadioButtonList ID="rbtList_MB_SNORE" runat="server" CssClass="CtlFnt" style="background:transparent" RepeatLayout="Flow" RepeatDirection="Horizontal" >
                        <asp:ListItem Text="是" Value="1" CssClass="CtlFnt"></asp:ListItem>
                        <asp:ListItem Text="否" Value="2" CssClass="CtlFnt"></asp:ListItem>
                        <asp:ListItem Text="不知道" Value="3" CssClass="CtlFnt"></asp:ListItem>
                    </asp:RadioButtonList>
                    (為安排住宿)									
				</div>
			</div>
			<div class="row" >
				<div class="col-md-2 thm">
					公司/學校系級
				</div>
				<div class="col-md-10">
					<asp:TextBox ID="SCHOOL" runat="server" class="form-control" />
				</div>
			</div>
			<div id="TR_G_8" runat="server" class="row" >
				<div class="col-md-2 thm">
					<span style="color:Red">**</span>
					身心狀況
				</div>
				<div class="col-md-10">
                    <asp:Repeater ID="dl_MB_SICK" runat="server">
                        <ItemTemplate>
                            <asp:HiddenField ID="hid_MB_SICK" runat="server" Value='<%#Container.DataItem("VALUE")%>' />
                            <asp:CheckBox ID="cb_MB_SICK" runat="server" Text='<%#Container.DataItem("TEXT")%>' CssClass="CtlFnt" />
                            <asp:TextBox ID="txt_MB_SICK" runat="server" Visible="false" class="form-control" ></asp:TextBox>
                        </ItemTemplate>
                    </asp:Repeater>
				</div>
			</div>
			<div id="TR_G_9" runat="server" class="row" >
				<div class="col-md-12 thm">
					修持法門
				</div>
			</div>
			<div id="TR_G_10" runat="server" class="row" >
				<div class="col-md-12">
                    您曾經修持過毗婆舍那禪法嗎?<BR/>
                    <asp:RadioButton ID="rbt_MB_PIPOSHENA_Y" Text="是" runat="server" GroupName="MB_PIPOSHENA" CssClass="CtlFnt" />
                    <asp:RadioButton ID="rbt_MB_PIPOSHENA_N" Text="否" runat="server" GroupName="MB_PIPOSHENA" CssClass="CtlFnt" />
                    <BR/>指導老師：
                    <asp:TextBox ID="txt_MB_TEACH" runat="server" class="form-control" />							
				</div>
			</div>
			<div id="TR_G_11" runat="server" class="row" >
				<div class="col-md-2 thm">
					您目前的修持法門
				</div>
				<div class="col-md-10">
					<asp:TextBox ID="txt_MB_FAMENNIAN" runat="server" class="form-control" />
				</div>
			</div>
			<div id="TR_G_12" runat="server" class="row" >
				<div class="col-md-12">
                    您過去曾經參加過七天以上的禪修嗎?<br/>
                    <asp:RadioButton ID="rbt_MB_OVER7DAY_Y" Text="是" runat="server" GroupName="MB_OVER7DAY" CssClass="CtlFnt" />
                    <asp:RadioButton ID="rbt_MB_OVER7DAY_N" Text="否" runat="server" GroupName="MB_OVER7DAY" CssClass="CtlFnt" />
                    <br/>地點：<br/>
                    <asp:TextBox ID="txt_MB_PLACE" runat="server" class="form-control" />					
				</div>
			</div>
			<div id="TR_G_13" runat="server" class="row" >
				<div class="col-md-12 thm">
                    <span style="color:Red">**</span>
                    緊急聯絡人									
				</div>
			</div>
			<div id="TR_G_14" runat="server" class="row" >
				<div class="col-md-2 thm">
					法名/姓名
				</div>
				<div class="col-md-4">
					<asp:TextBox ID="txt_MB_EMGCONT" runat="server" class="form-control" />
				</div>
				<div class="col-md-2 thm">
					電話/手機
				</div>
				<div class="col-md-4">
					<asp:TextBox ID="txt_MB_CONTMOBIL" runat="server" class="form-control" />
				</div>
			</div>
			<div id="TR_G_15_T" runat="server" class="row" >
				<div class="col-md-12 thm">
					捐助物資/捐款
				</div>
			</div>
			<div id="TR_G_15" runat="server" class="row" >
				<div class="col-md-12">
					<asp:TextBox ID="MB_AMTMEMO" runat="server" class="form-control" />
				</div>
			</div>
		</div>

        <div id="tb_Page3_Button" runat="server" class="text-center" visible="false" >
            <asp:Button ID="btn_Cancel" runat="server" Text="取消報名" CssClass="btn btn-info" Visible="false" />
            &nbsp;&nbsp;&nbsp;
            <asp:Button ID="btn_Save" runat="server" Text="確定" CssClass="btn btn-info" />            		
        </div>

        <div class="text-left">
            說明
            <br/>
            *:為所有課程必填項目,未填者將無法報名
            <br/>
            **:七天以上(含)之禪修課程必填            		
        </div>

        <asp:HiddenField ID="hid_HaveMember" runat="server" Value="False" />
    </form>
</body>
</html>
