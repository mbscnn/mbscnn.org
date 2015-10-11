<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MBMnt_ItemCode_01_v01.aspx.vb" Inherits="MBSC.MBMnt_ItemCode_01_v01" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>下拉選單維護作業</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<meta http-equiv="cachecontrol" content="no-cache">
		<meta HTTP-EQUIV="Pragma" content="no-cache">
		<meta http-equiv="Expires" content="-1">
        <!-- #include virtual="~/inc/MBSCCSS.inc" -->
        <!--CSS-->
        <!-- #include virtual="~/inc/MBSCJS.inc" -->
        <!--JS-->
		<script language="JavaScript">
		function popAdvSearch(bt){
			var text = document.getElementById("tbSearch");
			var qryText = encodeURI(text.value);
			//ddlSubSys.selectedvalue
			if (text.value.length < 1){
				alert('請輸入搜尋字串');
			}else{
				var sUrl = window.location.pathname  + '?AdvSearch=yes&QueryText=' + qryText ;
				var vReturn = window.showModalDialog(sUrl ,self ,'dialogWidth:' + window.screen.AvailWidth + 'px; dialogHeight:' + window.screen.AvailHeight + 'px; center:yes;scroll:1;status:0;help:0;resizable:0'); 
				if ((vReturn != undefined)&&(vReturn.length > 0)){
					document.getElementById("hidSearchCodeID").value = vReturn;
					bt.click();
					return true;
				}else{
					return false;
				}					
			}
		}	
		
		function sureToEdit(level){
			if (Editing()) return false;
			
			var select = document.all['ddlItemClass' + level];
			var value = select.value;
			if ((value == undefined)||(value == '-1')){
				alert("請選擇");
				return false;
			}else{
				return true;
			}
		}
		
		function sureToAdd(){
			if (Editing()) return false;
			return true;
		}
		
		function sureToDel(level,bt){
			if (Editing()) return false;
			
			var select = document.all['ddlItemClass' + level];
			var hasSon = document.all['ddlItemClass' + level].getAttribute('hasSon');
			var value = select.value;
			if ((value == undefined)||(value == '-1')){
				alert("請選擇");
			}else if(hasSon == 'true'){
				alert('尚有子項目類別未刪除！');
			}else{
				if (confirm('確定刪除？')) bt.click();
			}
		}
		
		function Editing(){
			if (divEdit.style.display == ""){
				alert('請先取消目前編輯！');
				return true;
			}else{
				return false;
			}		
		}
		
		window.onbeforeunload = function(){
			if(event.clientX > (document.body.clientWidth-25) && event.clientY < 0 || event.altKey){
					parent.window.returnValue=""; 
			}				
		} 

		//Sure Button in the Popping Window
		function returnCodeID(){
			if((document.getElementsByName("rdoCodeID")!=undefined)&&(document.getElementsByName("rdoCodeID").length > 0)){
				var rdoList = document.getElementsByName("rdoCodeID");
				var returnValue='';
				for(var i= 0 ; i < rdoList.length ; i++){
					if (rdoList[i].checked==true){
						returnValue = rdoList[i].attributes("codeid").nodeValue;
						break;
					}
				}
				parent.window.returnValue=returnValue; 
			}
			window.close();
		}		
		</script>
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
            <!-- #include virtual="~/inc/PageTab.inc" -->
            <table width="100%" cellspacing="0" cellpadding="0" style="width:1235px;background:transparent;margin-left: auto; margin-right: auto;" align="center">
                <tr>
                    <td style="vertical-align:top;padding:0;text-align:left;background:transparent;width:1050px" >
                        <!-- #include virtual="~/inc/Signin.inc" -->
			            <div style="FONT-SIZE: 17pt; COLOR: #800080; FONT-FAMILY: 標楷體; TEXT-ALIGN: center">下拉選單維護作業</div>
			            <br/>
			            <!-- #include virtual="~/inc/MBSCErrorMsg.inc" --> <!--錯誤訊息區-->
			            <div id="divNormal" runat="server">
				            <!-- #include virtual="~/inc/MBSCTableStart.inc" --> <!--顯示上方外框線-->
				            <table width="100%" cellpadding="0" cellspacing="o" class="CRTable_Top">
					            <tr>
						            <td class="th1c_b" width="10%">系統類別
						            </td>
						            <td class="td2_b" width="25%">
                                        <asp:DropDownList id="ddlSubSys" Runat="server" Width="90px" AutoPostBack="True" />
						            </td>
						            <td class="th1c_b" width="10%">
                                        搜尋系統類別
						            </td>
						            <td class="td2_b" width="25%" onkeydown="return false;">
                                        <asp:TextBox id="tbSearch" Runat="server" Columns="20"></asp:TextBox>&nbsp;
							            <asp:Button ID="btnSearch" Runat="server" Text="搜尋" CssClass="bt" onMouseDown="JavaScript:popAdvSearch(this);"></asp:Button>
                                        <asp:DropDownList ID="DDL_LEVEL" runat="server">
                                            <asp:ListItem Text="0" Value="0"></asp:ListItem>
                                            <asp:ListItem Text="1" Value="1"></asp:ListItem>
                                            <asp:ListItem Text="2" Value="2"></asp:ListItem>
                                            <asp:ListItem Text="3" Value="3"></asp:ListItem>
                                            <asp:ListItem Text="4" Value="4"></asp:ListItem>
                                            <asp:ListItem Text="5" Value="5"></asp:ListItem>
                                        </asp:DropDownList>
						            </td>
                                    <td class="th1c_b" width="10%" onkeydown="return false;">
                                        中心別
                                    </td>
                                    <td class="td2_b" width="20%" onkeydown="return false;">
                                        <asp:DropDownList ID="DDL_BRID" runat="server" />
                                    </td>
					            </tr>
				            </table>
				            <table height="100%" width="100%" cellpadding="0" cellspacing="0" class="CRTable">
					            <tr>
						            <td class="th1c_b" width="20%" height="20">項目類別</td>
						            <td class="td2c_b" align="center" width="16%"><asp:DropDownList Visible="False" id="ddlItemClass0" Runat="server" Width="100%" AutoPostBack="True"
								            UpCode=""></asp:DropDownList></td>
						            <td class="td2c_b" align="center" width="16%"><asp:DropDownList Visible="False" id="ddlItemClass1" Runat="server" Width="100%" AutoPostBack="True"
								            UpCode=""></asp:DropDownList></td>
						            <td class="td2c_b" align="center" width="16%"><asp:DropDownList Visible="False" id="ddlItemClass2" Runat="server" Width="100%" AutoPostBack="True"
								            UpCode=""></asp:DropDownList></td>
						            <td class="td2c_b" align="center" width="16%"><asp:DropDownList Visible="False" id="ddlItemClass3" Runat="server" Width="100%" AutoPostBack="True"
								            UpCode=""></asp:DropDownList></td>
						            <td class="td2c_b" align="center" width="16%"><asp:DropDownList Visible="False" id="ddlItemClass4" Runat="server" Width="100%"></asp:DropDownList><!--最後一個不用PostBack--></td>
					            </tr>
					            <tr>
						            <td class="th1c_b" width="20%" height="20">新增/修改/刪除</td>
						            <td class="td2c_b" align="center" width="16%" id="tdLevel0" runat="server">
							            <asp:ImageButton Visible="False" Runat="server" ID="imgbtnAdd0" ImageUrl="~/img/imgAdd.gif" onMouseDown="JavaScript:sureToAdd();"></asp:ImageButton>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							            <asp:ImageButton Visible="False" Runat="server" ID="imgbtnEdit0" ImageUrl="~/img/imgEdit.gif" onMouseDown="JavaScript:sureToEdit(0);"></asp:ImageButton>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							            <asp:ImageButton Visible="False" Runat="server" ID="imgbtnDel0" ImageUrl="~/img/imgDelete.gif" onMouseDown="JavaScript:sureToDel(0,this);"></asp:ImageButton>
						            </td>
						            <td class="td2c_b" align="center" width="16%" id="tdLevel1" runat="server">
							            <asp:ImageButton Visible="False" Runat="server" ID="imgbtnAdd1" ImageUrl="~/img/imgAdd.gif" onMouseDown="JavaScript:sureToAdd();"></asp:ImageButton>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							            <asp:ImageButton Visible="False" Runat="server" ID="imgbtnEdit1" ImageUrl="~/img/imgEdit.gif" onMouseDown="JavaScript:sureToEdit(1);"></asp:ImageButton>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							            <asp:ImageButton Visible="False" Runat="server" ID="imgbtnDel1" ImageUrl="~/img/imgDelete.gif" onMouseDown="JavaScript:sureToDel(1,this);"></asp:ImageButton>
						            </td>
						            <td class="td2c_b" align="center" width="16%" id="tdLevel2" runat="server">
							            <asp:ImageButton Visible="False" Runat="server" ID="imgbtnAdd2" ImageUrl="~/img/imgAdd.gif" onMouseDown="JavaScript:sureToAdd();"></asp:ImageButton>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							            <asp:ImageButton Visible="False" Runat="server" ID="imgbtnEdit2" ImageUrl="~/img/imgEdit.gif" onMouseDown="JavaScript:sureToEdit(2);"></asp:ImageButton>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							            <asp:ImageButton Visible="False" Runat="server" ID="imgbtnDel2" ImageUrl="~/img/imgDelete.gif" onMouseDown="JavaScript:sureToDel(2,this);"></asp:ImageButton>
						            </td>
						            <td class="td2c_b" align="center" width="16%" id="tdLevel3" runat="server">
							            <asp:ImageButton Visible="False" Runat="server" ID="imgbtnAdd3" ImageUrl="~/img/imgAdd.gif" onMouseDown="JavaScript:sureToAdd();"></asp:ImageButton>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							            <asp:ImageButton Visible="False" Runat="server" ID="imgbtnEdit3" ImageUrl="~/img/imgEdit.gif" onMouseDown="JavaScript:sureToEdit(3);"></asp:ImageButton>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							            <asp:ImageButton Visible="False" Runat="server" ID="imgbtnDel3" ImageUrl="~/img/imgDelete.gif" onMouseDown="JavaScript:sureToDel(3,this);"></asp:ImageButton>
						            </td>
						            <td class="td2c_b" align="center" width="16%" id="tdLevel4" runat="server">
							            <asp:ImageButton Visible="False" Runat="server" ID="imgbtnAdd4" ImageUrl="~/img/imgAdd.gif" onMouseDown="JavaScript:sureToAdd();"></asp:ImageButton>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							            <asp:ImageButton Visible="False" Runat="server" ID="imgbtnEdit4" ImageUrl="~/img/imgEdit.gif" onMouseDown="JavaScript:sureToEdit(4);"></asp:ImageButton>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							            <asp:ImageButton Visible="False" Runat="server" ID="imgbtnDel4" ImageUrl="~/img/imgDelete.gif" onMouseDown="JavaScript:sureToDel(4,this);"></asp:ImageButton>
						            </td>
					            </tr>
				            </table>
				            <div id="divEdit" runat="server" style="DISPLAY:none">
					            <table width="100%" cellpadding="0" cellspacing="0" class="CRTable">
						            <tr>
							            <td class="th1c_b" width="5%">ID</td>
							            <td class="td2c_b" width="5%" align="center"><asp:Label ID="lblID" Runat="server"></asp:Label></td>
							            <td class="th1c_b" width="10%">值</td>
							            <td class="td2_b" width="15%"><asp:TextBox id="tbValue" Runat="server" Width="100%" TextMode="SingleLine" onkeydown="KCmaxWord(this,199);"></asp:TextBox></td>
							            <td width="20%" class="th1c_b">項目文字描述</td>
							            <td class="td2_b" width="45%"><asp:TextBox id="tbText" Runat="server" Width="100%" TextMode="SingleLine" onkeydown="KCmaxWord(this,499);"></asp:TextBox></td>
						            </tr>
					            </table>
					            <table width="100%" cellpadding="0" cellspacing="0" class="CRTable">
						            <tr>
							            <td width="5%" class="th1c_b" nowrap>UPCODE</td>
							            <td class="td2c_b" width="5%" align="center" nowrap><asp:Label ID="lblUPCODE" Runat="server"></asp:Label></td>
							            <td width="5%" class="th1c_b" nowrap>排位<br>
								            順序</td>
							            <td class="td2_b" width="5%"><asp:TextBox id="tbSeq" Runat="server" Columns="4" MaxLength="4" TextMode="SingleLine" onkeydown="return KCAmtCheck();"></asp:TextBox></td>
							            <td width="5%" class="th1c_b" nowrap>備註<br>
								            說明</td>
							            <td class="td2_b" width="65%"><asp:TextBox id="tbPs" Runat="server" Width="100%" TextMode="MultiLine" Rows="3" onkeydown="KCmaxWord(this,499);"></asp:TextBox></td>
							            <td width="5%" class="th1c_b" nowrap>顯示<br>
								            隱藏</td>
							            <td class="td2_b" width="5%">
								            <asp:DropDownList id="ddlVisible" Runat="server" Width="80px">
									            <asp:ListItem Selected="False" Value="1">顯示</asp:ListItem>
									            <asp:ListItem Value="0">隱藏</asp:ListItem>
								            </asp:DropDownList>
							            </td>
						            </tr>
					            </table>
					            <table width="100%" cellpadding="0" cellspacing="0" class="CRTable">
						            <tr>
							            <td class="td2_b" >
								            <asp:Button id="btnSure" Runat="server" Text="確定" CssClass="bt"></asp:Button>&nbsp;
								            <asp:Button id="btnCancel" Runat="server" Text="取消" CssClass="bt"></asp:Button>
							            </td>
						            </tr>
					            </table>
				            </div>
				            <!-- #include virtual="~/inc/MBSCTableEnd.inc"--> <!--顯示下方外框線-->
				            <input id="hidCurrentUpCode" type="hidden" runat="server" NAME="hidCurrentUpCode">
				            <input id="hidCurrentLevel" type="hidden" runat="server" NAME="hidCurrentLevel">
				            <input id="hidSearchCodeID" type="hidden" runat="server" NAME="hidSearchCodeID">
			            </div>
			            <div id="divPopUp" runat="server">
				            <asp:DataGrid ID="dg" Runat="server" Width="100%" AutoGenerateColumns="False" CellPadding="0"
					            BorderWidth="0" CellSpacing="0" CssClass="CRTable" ShowFooter="False">
					            <Columns>
						            <asp:TemplateColumn>
							            <HeaderStyle CssClass="th1c_b"></HeaderStyle>
							            <HeaderTemplate>
								            點選
							            </HeaderTemplate>
							            <ItemStyle CssClass="td2c_b" HorizontalAlign="center"></ItemStyle>
							            <ItemTemplate>
								            <input type="radio" id="rdoCodeID" name="rdoCodeID" codeid='<%#Container.DataItem("CODEID")%>' />
							            </ItemTemplate>
						            </asp:TemplateColumn>
						            <asp:TemplateColumn>
							            <HeaderStyle CssClass="th1c_b"></HeaderStyle>
							            <HeaderTemplate>
								            系統名稱
							            </HeaderTemplate>
							            <ItemStyle CssClass="td2c_b" HorizontalAlign="center"></ItemStyle>
							            <ItemTemplate>
								            <%#getSubSysName(Container.DataItem("SUBSYSID"))%>
							            </ItemTemplate>
						            </asp:TemplateColumn>
						            <asp:BoundColumn HeaderStyle-CssClass="th1c_b" DataField="UPCODE" HeaderText="父類別" ItemStyle-HorizontalAlign="Center"></asp:BoundColumn>
						            <asp:BoundColumn HeaderStyle-CssClass="th1c_b" DataField="VALUE" HeaderText="值"></asp:BoundColumn>
						            <asp:BoundColumn HeaderStyle-CssClass="th1c_b" DataField="TEXT" HeaderText="項目文字描述"></asp:BoundColumn>
						            <asp:BoundColumn HeaderStyle-CssClass="th1c_b" DataField="NOTE" HeaderText="備註說明"></asp:BoundColumn>
					            </Columns>
				            </asp:DataGrid>
				            <div style="TEXT-ALIGN:center">
					            <input type="button" class="bt" value="確定" onclick="JavaScript: returnCodeID();">&nbsp;&nbsp;&nbsp;
					            <input type="button" class="bt" value="取消" onclick="JavaScript: window.close();">
				            </div>
				            <script language="JavaScript">
				                //將第一個 Radio Button設為checked
				                if ((document.getElementsByName("rdoCodeID") != undefined) && (document.getElementsByName("rdoCodeID").length > 0)) {
				                    document.getElementsByName("rdoCodeID")[0].checked = true;
				                }
				            </script>
			            </div>
                    </td>
                </tr>
            </table>
		</form>
	</body>
</HTML>
