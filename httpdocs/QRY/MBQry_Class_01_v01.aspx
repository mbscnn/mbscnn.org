<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="MBQry_Class_01_v01.aspx.vb" Inherits="MBSC.MBQry_Class_01_v01" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>MBSC課程報名彙總表</title>
    <!-- #include virtual="~/inc/MBSCCSS.inc" -->
    <!-- #include virtual="~/inc/MBSCJS.inc" --> 
</head>
<body topmargin="0">
    <form id="form1" runat="server">
        <!-- #include virtual="~/inc/PageTab.inc" -->
        <!-- #include virtual="~/inc/Signin.inc" -->

        <!--錯誤訊息區-->
        <!-- #include virtual="~/inc/MBSCErrorMsg.inc" -->
        <div class="table-responsive" >
            <table class="table CRTable_Top" width="100%" cellspacing="0">
                <tr>
                    <td width="100%" class="mtrSecTitle_b" style="font-family: 標楷體;text-align:center" colspan="2">
                        <asp:Image ID="IMG_SET" runat="server" style="width: 32px; height: 32px;vertical-align:middle" />
                        報名中
                    </td>
                </tr>
                <asp:PlaceHolder ID="PLH_APLY_1_1" runat="server" >
                <tr>
                    <td width="15%" class="th1c_b">
                        課程地點
                    </td>
                    <td width="85%" class="td2_b">
                        <asp:DropDownList ID="DDL_MB_PLACE_1" runat="server" />
                    </td>
                </tr>
                <tr>
                    <td width="15%" class="th1c_b">
                        課程開始年月
                    </td>
                    <td width="85%" class="td2_b">
                        <asp:DropDownList ID="DDL_MB_SDATE_SY_1" runat="server" />年
                        <asp:DropDownList ID="DDL_MB_SDATE_SM_1" runat="server" />月至
                        <asp:DropDownList ID="DDL_MB_SDATE_EY_1" runat="server" />年
                        <asp:DropDownList ID="DDL_MB_SDATE_EM_1" runat="server" />月
                        <asp:Button ID="btnQRY1" runat="server" CssClass="bt" Text="確定" />  
                    </td>
                </tr>
                </asp:PlaceHolder>
            </table>
        </div>
        <div class="table-responsive" >
        <table class="table CRTable" width="100%" cellspacing="0">
            <asp:PlaceHolder ID="PLH_APLY_1_2" runat="server" >
            <tr>
                <td id="TD_1_1" runat="server" width="10%" class="th1c_b">
                    課程編號
                </td>
                <td id="TD_1_2" runat="server" width="10%" class="th1c_b">
                    地點
                </td>
                <td width="10%" class="th1c_b">
                    課程起訖日
                </td>
                <td width="15%" class="th1c_b">
                    課程名稱
                </td>
                <td id="TD_1_3" runat="server" width="10%" class="th1c_b">
                    指導老師
                </td>
                <td id="TD_1_4" runat="server" width="10%" class="th1c_b">
                    招收人數
                </td>
                <td  id="TD_1_5" runat="server" width="15%" class="th1c_b" style="display:none" >
                    報名起訖日
                </td>
                <td id="TD_1_6" runat="server" width="10%" class="th1c_b">
                    報名
                </td>
                <td id="TD_1_7" runat="server" width="10%" class="th1c_b">
                    發mail
                </td>
            </tr>
            </asp:PlaceHolder>
            <asp:Repeater ID="RP_CLASS_1" runat="server">
                <ItemTemplate>
                    <tr>
                        <td id="TD_1_1" runat="server" width="10%" class="td2Dc_b">
                            <!--課程編號-->
                            <asp:Literal ID="LTL_MB_SEQ" runat="server" Text='<%#Container.DataItem("MB_SEQ")%>' />    
                            <input type="hidden" id="MB_BATCH" runat="server" value='<%#Container.DataItem("MB_BATCH")%>' />                          
                        </td>
                        <td id="TD_1_2" runat="server" width="10%" class="td2Dc_b">
                            <!--地點-->
                            <%#Container.DataItem("MB_PLACE")%>
                        </td>
                        <td width="10%" class="td2Dc_b">
                            <!--課程起訖日-->
                            <asp:Literal ID="LTL_MB_SEDATE" runat="server" />
                        </td>
                        <td width="15%" class="td2Dc_b">
                            <!--課程名稱-->
                            <asp:LinkButton ID="MB_CLASS_NAME" runat="server" Text='<%#getMB_CLASS_NAME(Container.DataItem("MB_CLASS_NAME"), Container.DataItem("MB_BATCH"))%>' Style="text-decoration:underline" CommandName="CONTENT"  />
                        </td>
                        <td id="TD_1_3" runat="server" width="10%" class="td2Dc_b">
                            <!--指導老師-->
                            <%#Container.DataItem("MB_TEACHER")%>
                        </td>
                        <td id="TD_1_4" runat="server" width="10%" class="td2Dc_b">
                            <!--招收人數-->
                            <%#Container.DataItem("MB_FULL")%>
                        </td>
                        <td id="TD_1_5" runat="server" width="15%" class="td2Dc_b" style="display:none" >
                            <!--報名起訖日-->
                            <asp:Literal ID="LTL_MB_SEAPLY" runat="server" />
                        </td>
                        <td id="TD_1_6" runat="server" width="10%" class="td2Dc_b">
                            <!--報名-->
                            <asp:Button ID="btnSIGNUP" runat="server" CssClass="bt" Text="我要報名" CommandName="SIGNUP" />
                            <asp:Button ID="btnCANCEL" runat="server" CssClass="bt" Text="取消報名" CommandName="CANCEL" />
                            <asp:Label ID="LTL_APLY" runat="server" Style="font-size:14pt;font-weight:bold" Visible="false" Text="已截止報名" />
                        </td>
                        <td id="TD_1_7" runat="server" width="10%" class="td2Dc_b">
                            <!--發mail-->
                            <asp:Button ID="btnMB_ALERT1" runat="server" CssClass="bt" Text="提醒信一" CommandName="MB_ALERT1" Visible="false" />
                            <asp:Button ID="btnMB_ALERT2" runat="server" CssClass="bt" Text="提醒信二" CommandName="MB_ALERT2" Visible="false" />
                            <asp:Label ID="LTL_MB_ALERT2" runat="server" style="color:red;font-weight:bold" />
                        </td>
                    </tr>
                </ItemTemplate>
            </asp:Repeater>
        </table>
        </div>
        <BR/>
        <div class="table-responsive" >
            <table class="table CRTable_Top" width="100%" cellspacing="0">
                <tr>
                    <td width="100%" class="mtrSecTitle_b" style="font-family: 標楷體;text-align:center" colspan="2">
                        <asp:Image ID="IMG_PRE" runat="server" style="width: 32px; height: 32px;vertical-align:middle" />
                        課程活動預告
                    </td>                
                </tr>
                <asp:PlaceHolder ID="PLH_APLY_2_1" runat="server" >
                    <tr>
                        <td width="15%" class="th1c_b">
                            課程地點
                        </td>
                        <td width="85%" class="td2_b">
                            <asp:DropDownList ID="DDL_MB_PLACE_2" runat="server" />
                        </td>
                    </tr>
                    <tr>
                        <td width="15%" class="th1c_b">
                            課程開始年月
                        </td>
                        <td width="85%" class="td2_b">
                            <asp:DropDownList ID="DDL_MB_SDATE_SY_2" runat="server" />年
                            <asp:DropDownList ID="DDL_MB_SDATE_SM_2" runat="server" />月至
                            <asp:DropDownList ID="DDL_MB_SDATE_EY_2" runat="server" />年
                            <asp:DropDownList ID="DDL_MB_SDATE_EM_2" runat="server" />月
                            <asp:Button ID="btnQRY2" runat="server" CssClass="bt" Text="確定" />  
                        </td>
                    </tr>
                </asp:PlaceHolder>
            </table>
        </div>
        <div class="table-responsive" >
            <table class="table CRTable" width="100%" cellspacing="0">
                <asp:PlaceHolder ID="PLH_APLY_2_2" runat="server" >
                <tr>
                    <td id="TD_2_1" runat="server" width="10%" class="th1c_b">
                        課程編號
                    </td>
                    <td id="TD_2_2" runat="server" width="10%" class="th1c_b">
                        地點
                    </td>
                    <td width="20%" class="th1c_b">
                        課程起訖日
                    </td>
                    <td width="20%" class="th1c_b">
                        課程名稱
                    </td>
                    <td id="TD_2_3" runat="server" width="10%" class="th1c_b">
                        指導老師
                    </td>
                    <td id="TD_2_4" runat="server" width="10%" class="th1c_b">
                        招收人數
                    </td>
                    <td id="TD_2_5" runat="server" width="20%" class="th1c_b" style="display:none" >
                        報名起訖日
                    </td>
                </tr>
                </asp:PlaceHolder>
                <asp:Repeater ID="RP_CLASS_2" runat="server">
                    <ItemTemplate>
                        <tr>
                            <td id="TD_2_1" runat="server" width="10%" class="td2Dc_b">
                                <!--課程編號-->
                                <asp:Literal ID="LTL_MB_SEQ" runat="server" Text='<%#Container.DataItem("MB_SEQ")%>' />       
                            </td>
                            <td id="TD_2_2" runat="server" width="10%" class="td2Dc_b">
                                <!--地點-->
                                <%#Container.DataItem("MB_PLACE")%>
                            </td>
                            <td width="20%" class="td2Dc_b">
                                <!--課程起訖日-->
                                <asp:Literal ID="LTL_MB_SEDATE" runat="server" />
                            </td>
                            <td width="20%" class="td2Dc_b">
                                <!--課程名稱-->
                                <asp:LinkButton ID="MB_CLASS_NAME" runat="server" Text='<%#getMB_CLASS_NAME(Container.DataItem("MB_CLASS_NAME"), Container.DataItem("MB_BATCH"))%>' Style="text-decoration:underline" CommandName="CONTENT" />
                            </td>
                            <td id="TD_2_3" runat="server" width="10%" class="td2Dc_b">
                                <!--指導老師-->
                                <%#Container.DataItem("MB_TEACHER")%>
                            </td>
                            <td id="TD_2_4" runat="server" width="10%" class="td2Dc_b">
                                <!--招收人數-->
                                <%#Container.DataItem("MB_FULL")%>
                            </td>
                            <td id="TD_2_5" runat="server" width="20%" class="td2Dc_b" style="display:none" >
                                <!--報名起訖日-->
                                <asp:Literal ID="LTL_MB_SEAPLY" runat="server" />
                            </td>
                        </tr>
                    </ItemTemplate>
                </asp:Repeater>
            </table>
        </div>
        <BR/>
        <div class="table-responsive" >
            <table class="table CRTable_Top" width="100%" cellspacing="0">
                <tr>
                    <td width="100%" class="mtrSecTitle_b" style="font-family: 標楷體;text-align:center" colspan="2">
                        <asp:Image ID="IMG_RUN" runat="server" style="width: 32px; height: 32px;vertical-align:middle" />
                        進行中課程
                    </td>                
                </tr>
                <asp:PlaceHolder ID="PLH_APLY_3_1" runat="server" >
                <tr>
                    <td width="15%" class="th1c_b">
                        課程地點
                    </td>
                    <td width="85%" class="td2_b">
                        <asp:DropDownList ID="DDL_MB_PLACE_4" runat="server" />
                    </td>
                </tr>
                <tr>
                    <td width="15%" class="th1c_b">
                        課程開始年月
                    </td>
                    <td width="85%" class="td2_b">
                        <asp:DropDownList ID="DDL_MB_SDATE_SY_4" runat="server" />年
                        <asp:DropDownList ID="DDL_MB_SDATE_SM_4" runat="server" />月至
                        <asp:DropDownList ID="DDL_MB_SDATE_EY_4" runat="server" />年
                        <asp:DropDownList ID="DDL_MB_SDATE_EM_4" runat="server" />月
                        <asp:Button ID="btnQRY4" runat="server" CssClass="bt" Text="確定" />  
                    </td>
                </tr>
                </asp:PlaceHolder>
            </table>
        </div>
        <div class="table-responsive" >
            <table class="table CRTable" width="100%" cellspacing="0">
                <asp:PlaceHolder ID="PLH_APLY_3_2" runat="server" >
                <tr>
                    <td id="TD_3_1" runat="server" width="10%" class="th1c_b">
                        課程編號
                    </td>
                    <td id="TD_3_2" runat="server" width="10%" class="th1c_b">
                        地點
                    </td>
                    <td width="20%" class="th1c_b">
                        課程起訖日
                    </td>
                    <td width="20%" class="th1c_b">
                        課程名稱
                    </td>
                    <td id="TD_3_3" runat="server" width="10%" class="th1c_b">
                        指導老師
                    </td>
                    <td id="TD_3_4" runat="server" width="10%" class="th1c_b">
                        招收人數
                    </td>
                    <td id="TD_3_5" runat="server" width="20%" class="th1c_b" style="display:none" >
                        報名起訖日
                    </td>
                </tr>
                </asp:PlaceHolder>
                <asp:Repeater ID="RP_CLASS_4" runat="server">
                    <ItemTemplate>
                        <tr>
                            <td id="TD_3_1" runat="server" width="10%" class="td2Dc_b">
                                <!--課程編號-->
                                <asp:Literal ID="LTL_MB_SEQ" runat="server" Text='<%#Container.DataItem("MB_SEQ")%>' />       
                            </td>
                            <td id="TD_3_2" runat="server" width="10%" class="td2Dc_b">
                                <!--地點-->
                                <%#Container.DataItem("MB_PLACE")%>
                            </td>
                            <td width="20%" class="td2Dc_b">
                                <!--課程起訖日-->
                                <asp:Literal ID="LTL_MB_SEDATE" runat="server" />
                            </td>
                            <td width="20%" class="td2Dc_b">
                                <!--課程名稱-->
                                <asp:LinkButton ID="MB_CLASS_NAME" runat="server" Text='<%#getMB_CLASS_NAME(Container.DataItem("MB_CLASS_NAME"), Container.DataItem("MB_BATCH"))%>' Style="text-decoration:underline" CommandName="CONTENT" />
                            </td>
                            <td id="TD_3_3" runat="server" width="10%" class="td2Dc_b">
                                <!--指導老師-->
                                <%#Container.DataItem("MB_TEACHER")%>
                            </td>
                            <td id="TD_3_4" runat="server" width="10%" class="td2Dc_b">
                                <!--招收人數-->
                                <%#Container.DataItem("MB_FULL")%>
                            </td>
                            <td id="TD_3_5" runat="server" width="20%" class="td2Dc_b" style="display:none" >
                                <!--報名起訖日-->
                                <asp:Literal ID="LTL_MB_SEAPLY" runat="server" />
                            </td>
                        </tr>
                    </ItemTemplate>
                </asp:Repeater>
            </table>
        </div>
        <BR />
        <div class="table-responsive" >
            <table class="table CRTable_Top" width="100%" cellspacing="0">
                <tr>
                    <td width="100%" class="mtrSecTitle_b" style="font-family: 標楷體;text-align:center" colspan="2">
                        <asp:Image ID="IMG_NOSIGN" runat="server" style="width: 32px; height: 32px;vertical-align:middle" />
                        截止報名尚未開課
                    </td>                
                </tr>
                <asp:PlaceHolder ID="PLH_APLY_5_1" runat="server" >
                <tr>
                    <td width="15%" class="th1c_b">
                        課程地點
                    </td>
                    <td width="85%" class="td2_b">
                        <asp:DropDownList ID="DDL_MB_PLACE_5" runat="server" />
                    </td>
                </tr>
                <tr>
                    <td width="15%" class="th1c_b">
                        課程開始年月
                    </td>
                    <td width="85%" class="td2_b">
                        <asp:DropDownList ID="DDL_MB_SDATE_SY_5" runat="server" />年
                        <asp:DropDownList ID="DDL_MB_SDATE_SM_5" runat="server" />月至
                        <asp:DropDownList ID="DDL_MB_SDATE_EY_5" runat="server" />年
                        <asp:DropDownList ID="DDL_MB_SDATE_EM_5" runat="server" />月
                        <asp:Button ID="btnQRY5" runat="server" CssClass="bt" Text="確定" />  
                    </td>
                </tr>
                </asp:PlaceHolder>
            </table>
        </div>
        <div class="table-responsive" >
            <table class="table CRTable" width="100%" cellspacing="0">
                <asp:PlaceHolder ID="PLH_APLY_5_2" runat="server" >
                <tr>
                    <td id="TD_5_1" runat="server" width="10%" class="th1c_b">
                        課程編號
                    </td>
                    <td id="TD_5_2" runat="server" width="10%" class="th1c_b">
                        地點
                    </td>
                    <td width="15%" class="th1c_b">
                        課程起訖日
                    </td>
                    <td width="20%" class="th1c_b">
                        課程名稱
                    </td>
                    <td id="TD_5_3" runat="server" width="10%" class="th1c_b">
                        指導老師
                    </td>
                    <td id="TD_5_4" runat="server" width="10%" class="th1c_b">
                        招收人數
                    </td>
                    <td id="TD_5_5" runat="server" width="15%" class="th1c_b" style="display:none" >
                        報名起訖日
                    </td>
                    <td id="TD_5_6" runat="server" width="10%" class="th1c_b">
                        發mail
                    </td>
                </tr>
                </asp:PlaceHolder>
                <asp:Repeater ID="RP_CLASS_5" runat="server">
                    <ItemTemplate>
                        <tr>
                            <td id="TD_5_1" runat="server" width="10%" class="td2Dc_b">
                                <!--課程編號-->
                                <asp:Literal ID="LTL_MB_SEQ" runat="server" Text='<%#Container.DataItem("MB_SEQ")%>' />
                                <input type="hidden" id="MB_BATCH" runat="server" value='<%#Container.DataItem("MB_BATCH")%>' />        
                            </td>
                            <td id="TD_5_2" runat="server" width="10%" class="td2Dc_b">
                                <!--地點-->
                                <%#Container.DataItem("MB_PLACE")%>
                            </td>
                            <td width="15%" class="td2Dc_b">
                                <!--課程起訖日-->
                                <asp:Literal ID="LTL_MB_SEDATE" runat="server" />
                            </td>
                            <td width="20%" class="td2Dc_b">
                                <!--課程名稱-->
                                <asp:LinkButton ID="MB_CLASS_NAME" runat="server" Text='<%#getMB_CLASS_NAME(Container.DataItem("MB_CLASS_NAME"), Container.DataItem("MB_BATCH"))%>' Style="text-decoration:underline" CommandName="CONTENT" />
                            </td>
                            <td id="TD_5_3" runat="server" width="10%" class="td2Dc_b">
                                <!--指導老師-->
                                <%#Container.DataItem("MB_TEACHER")%>
                            </td>
                            <td id="TD_5_4" runat="server" width="10%" class="td2Dc_b">
                                <!--招收人數-->
                                <%#Container.DataItem("MB_FULL")%>
                            </td>
                            <td id="TD_5_5" runat="server" width="15%" class="td2Dc_b" style="display:none" >
                                <!--報名起訖日-->
                                <asp:Literal ID="LTL_MB_SEAPLY" runat="server" />
                            </td>
                            <td id="TD_5_6" runat="server" width="10%" class="td2Dc_b">
                                <!--發mail-->
                                <asp:Button ID="btnMB_ALERT1" runat="server" CssClass="bt" Text="提醒信一" CommandName="MB_ALERT1" Visible="false" />
                                <asp:Button ID="btnMB_ALERT2" runat="server" CssClass="bt" Text="提醒信二" CommandName="MB_ALERT2" Visible="false" />
                                <asp:Label ID="LTL_MB_ALERT2" runat="server" style="color:red;font-weight:bold" />
                            </td>
                        </tr>
                    </ItemTemplate>
                </asp:Repeater>
            </table>
        </div>

        <asp:PlaceHolder ID="PLH_4" runat="server">
        <BR />
        <div class="table-responsive" >
            <table class="table CRTable_Top" width="100%" cellspacing="0">
                <tr>
                    <td width="100%" class="mtrSecTitle_b" style="font-family: 標楷體;text-align:center" colspan="2">
                        <asp:Image ID="IMG_HIS" runat="server" style="width: 32px; height: 32px;vertical-align:middle" />
                        課程記錄(已完成之課程)
                    </td>                
                </tr>
                <tr>
                    <td width="15%" class="th1c_b">
                        課程地點
                    </td>
                    <td width="85%" class="td2_b">
                        <asp:DropDownList ID="DDL_MB_PLACE_3" runat="server" />
                    </td>
                </tr>
                <tr>
                    <td width="15%" class="th1c_b">
                        課程開始年月
                    </td>
                    <td width="85%" class="td2_b">
                        <asp:DropDownList ID="DDL_MB_SDATE_SY_3" runat="server" />年
                        <asp:DropDownList ID="DDL_MB_SDATE_SM_3" runat="server" />月至
                        <asp:DropDownList ID="DDL_MB_SDATE_EY_3" runat="server" />年
                        <asp:DropDownList ID="DDL_MB_SDATE_EM_3" runat="server" />月
                        <asp:Button ID="btnQRY3" runat="server" CssClass="bt" Text="確定" />  
                    </td>
                </tr>
            </table>
        </div>
        <div class="table-responsive" >
            <table class="table CRTable" width="100%" cellspacing="0">
                <tr>
                    <td width="10%" class="th1c_b">
                        課程編號
                    </td>
                    <td width="10%" class="th1c_b">
                        地點
                    </td>
                    <td width="20%" class="th1c_b">
                        課程起訖日
                    </td>
                    <td width="20%" class="th1c_b">
                        課程名稱
                    </td>
                    <td width="10%" class="th1c_b">
                        指導老師
                    </td>
                    <td width="10%" class="th1c_b">
                        招收人數
                    </td>
                    <td width="20%" class="th1c_b" style="display:none" >
                        報名起訖日
                    </td>
                </tr>
                <asp:Repeater ID="RP_CLASS_3" runat="server">
                    <ItemTemplate>
                        <tr>
                            <td width="10%" class="td2Dc_b">
                                <!--課程編號-->
                                <asp:Literal ID="LTL_MB_SEQ" runat="server" Text='<%#Container.DataItem("MB_SEQ")%>' />       
                            </td>
                            <td width="10%" class="td2Dc_b">
                                <!--地點-->
                                <%#Container.DataItem("MB_PLACE")%>
                            </td>
                            <td width="20%" class="td2Dc_b">
                                <!--課程起訖日-->
                                <asp:Literal ID="LTL_MB_SEDATE" runat="server" />
                            </td>
                            <td width="20%" class="td2Dc_b">
                                <!--課程名稱-->
                                <asp:LinkButton ID="MB_CLASS_NAME" runat="server" Text='<%#getMB_CLASS_NAME(Container.DataItem("MB_CLASS_NAME"), Container.DataItem("MB_BATCH"))%>' Style="text-decoration:underline" CommandName="CONTENT"  />
                            </td>
                            <td width="10%" class="td2Dc_b">
                                <!--指導老師-->
                                <%#Container.DataItem("MB_TEACHER")%>
                            </td>
                            <td width="10%" class="td2Dc_b">
                                <!--招收人數-->
                                <%#Container.DataItem("MB_FULL")%>
                            </td>
                            <td width="20%" class="td2Dc_b" style="display:none" >
                                <!--報名起訖日-->
                                <asp:Literal ID="LTL_MB_SEAPLY" runat="server" />
                            </td>
                        </tr>
                    </ItemTemplate>
                </asp:Repeater>
            </table>
        </div>
        </asp:PlaceHolder>
    </form>
</body>
</html>
