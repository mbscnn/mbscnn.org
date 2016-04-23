<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="MBMnt_Class_01_v01.aspx.vb" Inherits="MBSC.MBMnt_Class_01_v01" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head runat="server">
    <title></title>
    <!-- #include virtual="~/inc/MBSCCSS.inc" -->
    <!-- #include virtual="~/inc/MBSCJS.inc" -->
    <link href="<%=com.Azion.EloanUtility.UIUtility.getRootPathClient() + "/js/jquery-ui-1.11.4.custom/jquery-ui.css"%>" rel="stylesheet" type="text/css" />
    <script src="<%=com.Azion.EloanUtility.UIUtility.getRootPathClient() + "/js/jquery-ui-1.11.4.custom/jquery-ui.js"%>" type="text/javascript"></script>
    <script language="javascript" type="text/javascript">
        function doCheck(thisObj) {
            if (!confirm('您確定要刪除此筆資料？')) {
                return;
            } else {
                //window.event.srcElement.click();
                thisObj.click();
            }
        }
        function chkOne(rdBtnID) {
            var rduser = $(document.getElementById(rdBtnID));
            rduser.closest('TR').addClass('SelectedRowStyle');
            rduser.checked = true;
            var list = rduser.closest('table').find("INPUT[type='radio']").not(rduser);
            list.attr('checked', false);
            list.closest('TR').removeClass('SelectedRowStyle');
        }

        $(function () {
            $("#EARLYDATE").datepicker({
                dateFormat: "yy/mm/dd"
            });
        });
    </script>
    <base target="_self" />
</head>
<body>
    <form id="form1" runat="server">
    <!-- #include virtual="~/inc/PageTab.inc" -->

<table width="100%" cellspacing="0" cellpadding="0" style="width:1235px;background:transparent;margin-left: auto; margin-right: auto;" align="center">
    <tr>
        <td style="vertical-align:top;padding:0;text-align:left;background:transparent;width:1050px" >
            <!-- #include virtual="~/inc/Signin.inc" -->
            <div align="center">
                <h2 style="color: #800080; font-family: 標楷體">
                    MBSC課程表維護
                </h2>
            </div>

            <!--顯示上方外框線-->
            <!-- #include virtual="~/inc/MBSCTableStart.inc" -->

            <!--錯誤訊息區-->
            <!-- #include virtual="~/inc/MBSCErrorMsg.inc" -->
            <table id="tbl_QRY" runat="server" class="CRTable_Top" width="100%" cellspacing="0">
                <tr>
                    <td width="15%" class="th1c_b">
                      課程類別
                    </td>
                    <td width="35%" class="td2_b">
                        <asp:DropDownList ID="ddl_TYPE" runat="server"></asp:DropDownList>
                    </td>
                    <td width="15%" class="th1c_b">
                        課程地點
                    </td>
                    <td width="35%" class="td2_b">
                        <asp:DropDownList ID="ddl_PLACE" runat="server"></asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td width="15%" class="th1c_b">
                        課程開始年月
                    </td>
                    <td width="85%" class="td2_b" colspan="3">
                        <asp:DropDownList ID="ddl_StartYear" runat="server">
                        </asp:DropDownList>年
                        <asp:DropDownList ID="ddl_StartMonth" runat="server">
                        </asp:DropDownList>月至
                        <asp:DropDownList ID="ddl_EndYear" runat="server">
                        </asp:DropDownList>年
                        <asp:DropDownList ID="ddl_EndMonth" runat="server">
                        </asp:DropDownList>月
                    </td>
                </tr>
                <tr>
                    <td class="th1c_b" align="center" colspan="4">
                        <asp:Button ID="btn_Confirm" runat="server" Text="查詢" CssClass="bt" />
                    </td>
                </tr>
            </table>
            <table id="tbl_A" runat="server" class="CRTable" width="100%" cellspacing="0">
                <tr>
                    <td width="15%" class="th1c_b">
                        <span style="color: red; font-weight: bold; font-size: medium">*</span>類別
                    </td>
                    <td width="35%" class="td2_b">
                        <asp:DropDownList ID="ddl_EditTYPE" runat="server" />
                    </td>
                    <td width="15%" class="th1c_b">
                        <span style="color: red; font-weight: bold; font-size: medium">*</span>梯次
                    </td>
                    <td width="35%" class="td2_b">
                        <asp:TextBox ID="txt_BATCH" runat="server" CssClass="bordertxt" width="60%" />
                        <span style="color:red;font-weight:bold;font-size:10pt">若無梯次請輸入0</span>
                    </td>
                </tr>
                <tr>
                    <td width="15%" class="th1c_b">
                        <span style="color: red; font-weight: bold; font-size: medium">*</span>地點
                    </td>
                    <td width="35%" class="td2_b">
                        <asp:DropDownList ID="ddl_EditPLACE" runat="server" AutoPostBack="true" ></asp:DropDownList>
                    </td>
                    <td width="15%" class="th1c_b">
                        <span style="color: red; font-weight: bold; font-size: medium">*</span>課程名稱
                    </td>
                    <td width="35%" class="td2_b">
                        <asp:TextBox ID="txt_ClassName" runat="server" CssClass="bordertxt" width="60%"/>
                    </td>
                </tr>
                <tr>
                    <td width="15%" class="th1c_b">
                        <span style="color: red; font-weight: bold; font-size: medium">*</span>課程時間
                    </td>
                    <td width="85%" class="td2_b" colspan="3">
                        <asp:DropDownList ID="ddl_EditStartYear" runat="server">
                        </asp:DropDownList>年
                        <asp:DropDownList ID="ddl_EditStartMonth" runat="server">
                        </asp:DropDownList>月
                        <asp:DropDownList ID="ddl_EditStartDay" runat="server">
                        </asp:DropDownList>日  起  迄
                        <asp:DropDownList ID="ddl_EditEndYear" runat="server">
                        </asp:DropDownList>年
                        <asp:DropDownList ID="ddl_EditEndMonth" runat="server">
                        </asp:DropDownList>月
                        <asp:DropDownList ID="ddl_EditEndDay" runat="server">
                        </asp:DropDownList>日  止
                    </td>
                </tr>
                <tr>
                    <td width="15%" class="th1c_b">
                        <span style="color: red; font-weight: bold; font-size: medium">*</span>課程報名起訖
                    </td>
                    <td width="85%" class="td2_b" colspan="3">
                        <asp:DropDownList ID="DDL_MB_SAPLY_Y" runat="server">
                        </asp:DropDownList>年
                        <asp:DropDownList ID="DDL_MB_SAPLY_M" runat="server">
                        </asp:DropDownList>月
                        <asp:DropDownList ID="DDL_MB_SAPLY_D" runat="server">
                        </asp:DropDownList>日  起  迄
                        <asp:DropDownList ID="DDL_MB_EAPLY_Y" runat="server">
                        </asp:DropDownList>年
                        <asp:DropDownList ID="DDL_MB_EAPLY_M" runat="server">
                        </asp:DropDownList>月
                        <asp:DropDownList ID="DDL_MB_EAPLY_D" runat="server">
                        </asp:DropDownList>日  止
                    </td>
                </tr>
                <tr>
                    <td width="15%" class="th1c_b">
                        指導老師
                    </td>
                    <td width="35%" class="td2_b">
                        <asp:TextBox ID="txt_TEACHER" runat="server" CssClass="bordertxt" width="60%"/>
                    </td>
                    <td width="15%" class="th1c_b">
                        <span style="color: red; font-weight: bold; font-size: medium">*</span>是否需核准
                    </td>
                    <td width="35%" class="td2_b">
                        <asp:RadioButtonList ID="rbt_APV" runat="server" RepeatDirection="Horizontal" AutoPostBack="true" >
                            <asp:ListItem Text="需核准不繳費" Value="1" Style="font-size:14pt"></asp:ListItem>
                            <asp:ListItem Text="不需核准" Value="2" Style="font-size:14pt"></asp:ListItem>
                            <asp:ListItem Text="需核准且要繳費" Value="3" Style="font-size:14pt"></asp:ListItem>
                        </asp:RadioButtonList>
                    </td>
                </tr>
                <tr>
                    <td width="15%" class="th1c_b">
                        額滿人數
                    </td>
                    <td width="35%" class="td2_b">
                       <asp:TextBox ID="txt_FULL" runat="server" CssClass="bordext" width="60%"/>人
                    </td>
                    <td width="15%" class="th1c_b">
                        備取人數
                    </td>
                    <td width="35%" class="td2_b">
                        <asp:TextBox ID="txt_WAIT" runat="server" CssClass="bordext" width="60%"/>人
                    </td>
                </tr>
                <tr>
                    <td width="15%" class="th1c_b">
                        是否開課
                    </td>
                    <td width="35%" class="td2_b">
                       <asp:RadioButtonList ID="rbt_YES" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow" >
                            <asp:ListItem Text="是" Value="Y" Style="font-size:14pt"></asp:ListItem>
                            <asp:ListItem Text="否" Value="N" Style="font-size:14pt"></asp:ListItem>
                        </asp:RadioButtonList>
                    </td>
                    <td width="15%" class="th1c_b">
                        課程編號
                    </td>
                    <td width="35%" class="td2_b">
                        <asp:Label ID="lbl_sSEQ" runat="server" Text=""></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td width="15%" class="th1c_b">
                        備註說明
                    </td>
                    <td width="85%" class="td2_b" colspan="3">
                       <asp:TextBox ID="txt_MEMO" runat="server" CssClass="bordext" width="85%"/>
                    </td>
                </tr>
                <tr>
                    <td width="15%" class="th1c_b">
                        <span style="color: red; font-weight: bold; font-size: medium">*</span>取消報名期限
                    </td>
                    <td width="35%" class="td2_b">
                        開課日
                        <asp:TextBox ID="MB_CDAYS" runat="server" Columns="3" Text="3" />
                        天前
                    </td>
                    <td width="15%" class="th1c_b">
                        <span style="color: red; font-weight: bold; font-size: medium">*</span>提醒信期限一/二
                    </td>
                    <td width="35%" class="td2_b">
                        開課日
                        <asp:TextBox ID="MB_ALERT1_DAY" runat="server" Columns="2" Text="10" />
                        天前及
                        <asp:TextBox ID="MB_ALERT2_DAY" runat="server" Columns="2" text="3" />
                        天前
                    </td>
                </tr>
                <tr>
                    <td width="15%" class="th1c_b">
                        <span style="color: red; font-weight: bold; font-size: medium">*</span>報到時間
                    </td>
                    <td width="35%" class="td2_b">
                        <asp:DropDownList ID="REGTIME_H" runat="server" />
                        :
                        <asp:TextBox ID="REGTIME_M" runat="server" Columns="2" />
                    </td>
                    <td width="15%" class="th1c_b">
                        <span style="color: red; font-weight: bold; font-size: medium">*</span>聯絡人/電話
                    </td>
                    <td width="35%" class="td2_b">
                        <asp:TextBox ID="CONTACT" runat="server" Columns="10" />
                        /
                        <asp:TextBox ID="CONTEL" runat="server" Columns="15" />
                    </td>
                </tr>
                <tr>
                    <td width="15%" class="th1c_b">
                        <span style="color: red; font-weight: bold; font-size: medium">*</span>上課地點
                    </td>
                    <td width="85%" class="td2_b" colspan="3">
                        <asp:DropDownList ID="DDL_CLASS_PLACE" runat="server" />
                    </td>
                </tr>
                <tr>
                    <td width="15%" class="th1c_b">
                        <span style="color: red; font-weight: bold; font-size: medium">*</span>交通資訊說明
                    </td>
                    <td width="85%" class="td2_b" colspan="3">
                        <asp:TextBox ID="TRAFFIC_DESC" runat="server" width="99%" />
                    </td>
                </tr>
                <tr>
                    <td width="15%" class="th1c_b">
                        <span style="color: red; font-weight: bold; font-size: medium">*</span>是否需填初學者
                    </td>
                    <td width="85%" class="td2_b" colspan="3">
                        <asp:RadioButtonList ID="MB_BEGIN" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow" >
                            <asp:ListItem Text="是" Value="Y" Style="font-size:14pt"  />
                            <asp:ListItem Text="否" Value="N" Style="font-size:14pt" />
                        </asp:RadioButtonList>
                    </td>
                </tr>
                <tr id="TR_CHARGE_1" runat="server" visible="false">
                    <td width="15%" class="th1c_b">
                        課程費用
                    </td>
                    <td width="85%" class="td2_b" colspan="3">
                        <asp:TextBox ID="TXT_CHARGE" runat="server" Columns="7" MaxLength="7" />
                    </td>
                </tr>
                <tr id="TR_CHARGE_2" runat="server" visible="false">
                    <td width="15%" class="th1c_b">
                        早鳥日期
                    </td>
                    <td width="35%" class="td2_b">
                        <asp:TextBox ID="EARLYDATE" runat="server" Columns="10" />
                    </td>
                    <td width="15%" class="th1c_b">
                        優惠費用
                    </td>
                    <td width="35%" class="td2_b">
                        <asp:TextBox ID="TXT_FAVCHARGE" runat="server" Columns="7" MaxLength="7" />
                    </td>
                </tr>
                <tr>
                    <td width="15%" class="th1c_b">
                        <span style="color: red; font-weight: bold; font-size: medium">*</span>
                        注意事項說明
                        <asp:Button ID="btnPREC" runat="server" CssClass="bt" Text="挑選預設詞彙" />
                    </td>
                    <td width="85%" class="td2_b" colspan="3">
                        <asp:TextBox ID="TXT_MB_PREC_MEMO" runat="server" TextMode="MultiLine" Width="99%" MaxLength="2000" Rows="3" />
                        <asp:PlaceHolder ID="PLH_PREC" runat="server" Visible="false">
                        <asp:Repeater ID="RP_PREC" runat="server" >
                            <ItemTemplate>
                                <div>
                                    <asp:CheckBox ID="CHB_PREC" runat="server" />
                                    <asp:Literal ID="LTL_PREC" runat="server" Text='<%#Container.DataItem("TEXT")%>' />
                                </div>
                            </ItemTemplate>
                        </asp:Repeater>
                        <asp:Button ID="btnPREC_YES" runat="server" CssClass="bt" Text="確定挑選" />
                        </asp:PlaceHolder>
                    </td>
                </tr>
                <tr runat="server" id="tr_Check">
                    <td colspan="4" class="th1c_b" align="center">
                        <asp:Button ID="btn_Add" runat="server" Text="新增課程" CssClass="bt" />
                        <asp:Button ID="btn_Batch" runat="server" Text="新增梯次" CssClass="bt" Visible="false" />
                        <asp:Button ID="btn_Edit" runat="server" Text="修改明細" CssClass="bt" Visible="false"/>
                        <asp:Button ID="btn_Type" runat="server" Text="修改課程名稱/類別" CssClass="bt" Visible="false" />
                    </td>
                </tr>
            </table>
            <asp:DataGrid ID="dg_Class" runat="server" CssClass="CRTABLE" Width="100%" CellSpacing="0" AutoGenerateColumns="False"  >
                    <Columns>
                        <asp:TemplateColumn HeaderText="挑選" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle CssClass="th1c_b" Wrap="False" Width="5%" />
                            <ItemStyle Wrap="True" CssClass="td2c_b"></ItemStyle>
                            <ItemTemplate>
                                <asp:RadioButton ID="rbt_SEQ" runat="server" onclick='chkOne(this.id)'></asp:RadioButton>
                            </ItemTemplate>
                        </asp:TemplateColumn>
                    </Columns>
                    <Columns>
                        <asp:TemplateColumn HeaderText="課程編號" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle CssClass="th1c_b" Wrap="False" Width="5%" />
                            <ItemStyle Wrap="True" CssClass="td2c_b"></ItemStyle>
                            <ItemTemplate>
                                <asp:Label ID="lbl_SEQ" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"MB_SEQ")%>'> ></asp:Label>
                            </ItemTemplate>
                        </asp:TemplateColumn>
                    </Columns>
                    <Columns>
                        <asp:TemplateColumn HeaderText="類別" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle CssClass="th1c_b" Wrap="False" Width="7%" />
                            <ItemStyle Wrap="True" CssClass="td2c_b"></ItemStyle>
                            <ItemTemplate>
                                <asp:Label ID="lbl_TYPE" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"MB_TYPE")%>'> ></asp:Label>
                            </ItemTemplate>
                        </asp:TemplateColumn>
                    </Columns>
                    <Columns>
                        <asp:TemplateColumn HeaderText="梯次" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle CssClass="th1c_b" Wrap="False" Width="5%" />
                            <ItemStyle Wrap="True" CssClass="td2c_b"></ItemStyle>
                            <ItemTemplate>
                                <input type="hidden" id="HID_MB_BATCH" runat="server" value='<%#Container.DataItem("MB_BATCH")%>' />
                                <asp:Label ID="lbl_BATCH" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"MB_BATCH")%>'> ></asp:Label>
                            </ItemTemplate>
                        </asp:TemplateColumn>
                    </Columns>
                    <Columns>
                        <asp:TemplateColumn HeaderText="地點" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle CssClass="th1c_b" Wrap="False" Width="7%" />
                            <ItemStyle Wrap="True" CssClass="td2c_b"></ItemStyle>
                            <ItemTemplate>
                                <asp:Label ID="lbl_PLACE" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"MB_PLACE")%>'> ></asp:Label>
                            </ItemTemplate>
                        </asp:TemplateColumn>
                    </Columns>
                    <Columns>
                        <asp:TemplateColumn HeaderText="課程起訖日" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle CssClass="th1c_b" Wrap="False" Width="15%" />
                            <ItemStyle Wrap="True" CssClass="td2c_b"></ItemStyle>
                            <ItemTemplate>
                                <asp:Label ID="lbl_SDATE" runat="server" Text='<%#ChangeDate(DataBinder.Eval(Container.DataItem,"MB_SDATE"))%>' ></asp:Label>
                                <asp:Label ID="lbl_EDATE" runat="server" Text='<%#ChangeDate(DataBinder.Eval(Container.DataItem,"MB_EDATE"))%>' ></asp:Label>
                            </ItemTemplate>
                        </asp:TemplateColumn>
                    </Columns>
                    <Columns>
                        <asp:TemplateColumn HeaderText="課程名稱" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle CssClass="th1c_b" Wrap="False" Width="15%" />
                            <ItemStyle Wrap="True" CssClass="td2c_b"></ItemStyle>
                            <ItemTemplate>
                                <asp:Label ID="lbl_CLASS_NAME" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"MB_CLASS_NAME")%>'> ></asp:Label>
                            </ItemTemplate>
                        </asp:TemplateColumn>
                    </Columns>
                    <Columns>
                        <asp:TemplateColumn HeaderText="指導老師" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle CssClass="th1c_b" Wrap="False" Width="5%" />
                            <ItemStyle Wrap="True" CssClass="td2c_b"></ItemStyle>
                            <ItemTemplate>
                                <asp:Label ID="lbl_TEACHER" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"MB_TEACHER")%>'> >
                                </asp:Label>
                            </ItemTemplate>
                        </asp:TemplateColumn>
                    </Columns>
                    <Columns>
                        <asp:TemplateColumn HeaderText="招收人數" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle CssClass="th1c_b" Wrap="False" Width="5%" />
                            <ItemStyle Wrap="True" CssClass="td2c_b"></ItemStyle>
                            <ItemTemplate>
                                <asp:Label ID="lbl_People" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"MB_FULL")%>'> ></asp:Label>
                            </ItemTemplate>
                        </asp:TemplateColumn>
                    </Columns>
                    <Columns>
                        <asp:TemplateColumn HeaderText="是否需核准 " HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle CssClass="th1c_b" Wrap="False" Width="5%" />
                            <ItemStyle Wrap="True" CssClass="td2c_b"></ItemStyle>
                            <ItemTemplate>
                                <asp:Label ID="lbl_APV" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"MB_APV")%>'> ></asp:Label>
                            </ItemTemplate>
                        </asp:TemplateColumn>
                    </Columns>
                    <Columns>
                        <asp:TemplateColumn HeaderText="修改" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle CssClass="th1c_b" Wrap="False" Width="3%" />
                            <ItemStyle Wrap="True" CssClass="td2c_b"></ItemStyle>
                            <ItemTemplate>
                                <asp:ImageButton ID="img_Edit" runat="server" OnClick="img_Edit_Click" ImageUrl="~/img/imgEdit.gif"/>
                            </ItemTemplate>
                        </asp:TemplateColumn>
                    </Columns>
                    <Columns>
                        <asp:TemplateColumn HeaderText="刪除" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle CssClass="th1c_b" Wrap="False" Width="3%" />
                            <ItemStyle Wrap="True" CssClass="td2c_b"></ItemStyle>
                            <ItemTemplate>
                                <asp:Label ID="lbl_OPEN" runat="server" Visible="false" Text='<%#DataBinder.Eval(Container.DataItem,"MB_OPEN")%>'> ></asp:Label>
                                <asp:ImageButton ID="img_Del" runat="server" OnClick="img_Del_Click" onmousedown="doCheck(this);" ImageUrl="~/img/imgDelete.gif"/>
                            </ItemTemplate>
                        </asp:TemplateColumn>
                    </Columns>
                </asp:DataGrid>
            <div align="center" >
                    <asp:Button ID="btn_QSEQ" runat="server" Text="確定" CssClass="bt" Visible="false" />
                </div>
            <!--顯示下方外框線-->
            <!--#include virtual="~/inc/MBSCTableEnd.inc"-->
            <div style="font-size:12pt;font-weight:bold;color:red">
                備註:<BR/>
                新增課程:是新開的課程,還未有課程編號<BR/>
                新增梯次:已經有課程編號請點原有課程可進行新增梯次
            </div>
        </td>
    </tr>
</table>
    </form>
</body>
</html>
