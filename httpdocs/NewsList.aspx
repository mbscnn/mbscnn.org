<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="NewsList.aspx.vb" Inherits="MBSC.NewsList" ValidateRequest="false" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<%@ Register Assembly="CKEditor.NET" Namespace="CKEditor.NET" TagPrefix="CKEditor" %>

<html>
<head>
    <title>MBSC佛陀原始正法中心</title>
    <!-- #include virtual="~/inc/MBSCCSS.inc" -->
    <!-- #include virtual="~/inc/MBSCJS.inc" -->
    <LINK href="<%=com.Azion.EloanUtility.UIUtility.getRootPath() + "/css/jquery.bxslider.css"%>" type="text/css" rel="StyleSheet" media="screen" >
    <script language="JavaScript" type="text/javascript" src="<%=com.Azion.EloanUtility.UIUtility.getRootPath() + "/ckeditor/ckeditor.js"%>"></script>
    <script language="javascript" type="text/javascript">
        $(document).ready(function () {
            //alert($("#RB_CLASS_YES").prop("checked"));
            if ($("#RB_CLASS_YES").prop("checked"))
            {
                $("#btnChoose").show();
            }
            else
            {
                $("#btnChoose").hide();
            }
        })

        function chooseClass(sPath)
        {
            //document.all("HID_MB_SEQ").value = "";
            //$("#HID_MB_SEQ").val("");
            //var strFeatures = 'dialogWidth=' + window.screen.width + 'px;dialogHeight=' + window.screen.height + 'px;center=yes;help=no;status=no;resizable=no';
            //var dialogAnswer = window.showModalDialog(sPath, self, strFeatures);

            var strFeatures = 'width=' + window.screen.width + 'px,height=' + window.screen.height + 'px,top=0, left=0, toolbar=yes, menubar=yes, scrollbars=yes, resizable=no,location=n o, status=no';
            var dialogAnswer = window.open(sPath, self, strFeatures);
            //if (dialogAnswer!=undefined)
            //{
                //document.all("HID_MB_SEQ").value = dialogAnswer;
            //    $("#HID_MB_SEQ").val(dialogAnswer);
            //}
        }

        function getRootPath() {
            var strFullPath = window.document.location.href;
            var strPath = window.document.location.pathname;
            var pos = strFullPath.indexOf(strPath);
            var prePath = strFullPath.substring(0, pos);
            var postPath = strPath.substring(0, strPath.substr(1).indexOf('/') + 1);
            return (prePath + postPath);
        }

        //CKEDITOR.on('dialogDefinition', function (ev) {
        //    // Take the dialog name and its definition from the event data
        //    var dialogName = ev.data.name;
        //    var dialogDefinition = ev.data.definition;
        //    if (dialogName == 'image') {
        //        dialogDefinition.onOk = function (e) {
        //            var imageSrcUrl = e.sender.originalElement.$.src;
        //            //e.sender.imageElement.$.outerHTML
        //            //<IMG style="BORDER-TOP: 1px solid; HEIGHT: 403px; BORDER-RIGHT: 1px solid; WIDTH: 403px; BORDER-BOTTOM: 1px solid; FLOAT: right; MARGIN-LEFT: 2px; BORDER-LEFT: 1px solid; MARGIN-RIGHT: 2px">
        //            var imgHtml = CKEDITOR.dom.element.createFromHtml("<img src=" + imageSrcUrl + " alt='' align='right' style='width:100px;height:100px' />");
        //            //CKEDITOR.instances["CKEditorControl2"].insertElement(imgHtml);
        //        };
        //    }
        //});
    </script>
    <style type="text/css">
        .NEWSTitle
        {
            font-family:"微軟正黑體","新細明體";
            font-size:16pt;
            font-weight:bold;
            text-align:center;
            padding:2px;
            color:white;            
            /*background:transparent;*/
            background: #003040;
            background: linear-gradient(#003040, #002535);

        }
        .NEWSTh1c
        {
            font-family:"微軟正黑體","新細明體";font-size:14pt;font-weight:bold;text-align:center;padding:4px;background:transparent
        }
        .NEWSTd2c
        {
            font-family:"微軟正黑體","新細明體";font-size:12pt;text-align:center;padding:4px;background:transparent
        }
    </style>
</head>
<body topmargin="0" >
    <form id="form1" runat="server">
        <!-- #include virtual="~/inc/PageTab.inc" -->
        <!--錯誤訊息區-->
        <!-- #include virtual="~/inc/MBSCErrorMsg.inc" -->
        <table width="100%" cellspacing="0" cellpadding="0" style="width:1235px;background:transparent;margin-left: auto; margin-right: auto;" align="center">
            <tr>
                <td style="vertical-align:top;padding:0;text-align:left;background:transparent;width:185px;" >
                    <!-- #include virtual="~/inc/vTab.inc" -->
                </td>
                <td id="TD_C" style="vertical-align:top;padding:0;text-align:left;background:transparent;width:1050px;" >
                    <!-- #include virtual="~/inc/Signin.inc" -->
                    <asp:PlaceHolder ID="PLH_CLASS_LIST" runat="server" >
                    <table width="100%" cellspacing="0" style="padding:0;background:transparent;">
                        <tr>
                            <td style="width:640px;vertical-align:top;background:transparent;">                                
                                <ul class="bxslider">
                                  <asp:Repeater ID="RP_Banner" runat="server">
                                      <ItemTemplate>
                                          <li>
                                              <img id="IMG_BANNER" runat="server" />
                                          </li>
                                      </ItemTemplate>
                                  </asp:Repeater>
                                </ul>
                            </td>
                            <td style="width:410px;vertical-align:top;background:transparent;" >
                                <table border="0" borderColor="black"  style="width:100%;border-collapse:collapse;empty-cells:show;background:transparent;" cellspacing="0" cellpadding="0" align="left"  >
                                    <tr>
                                        <td width="100%" class="NEWSTitle" colspan="2">
                                            報名中
                                        </td>
                                    </tr>
                                    <asp:Repeater ID="RP_CLASS_1" runat="server">
                                        <ItemTemplate>
                                            <tr>
                                                <td width="50%" class="NEWSTd2c">
                                                    <!--課程名稱-->
                                                    <asp:LinkButton ID="LinkButton2" runat="server" Text='<%#Container.DataItem("MB_CLASS_NAME")%>' CommandArgument='<%#Container.DataItem("MB_SEQ")%>' Style="text-decoration:underline" CommandName="CONTENT" />                                                     
                                                </td>
                                            </tr>
                                        </ItemTemplate>
                                    </asp:Repeater>

                                    <tr>
                                        <td width="100%" class="NEWSTitle" colspan="2">
                                            <asp:LinkButton ID="LB_CLASS_4" runat="server" Text="進行中課程" ForeColor="White" style="text-decoration:underline" />
                                        </td>                
                                    </tr>
                                    <asp:PlaceHolder ID="PLH_CLASS_4" runat="server" Visible="false" >
                                    <asp:Repeater ID="RP_CLASS_4" runat="server">
                                        <ItemTemplate>
                                            <tr>
                                                <td width="50%" class="NEWSTd2c">
                                                    <!--課程名稱-->
                                                    <asp:LinkButton ID="MB_SEQ" runat="server" Text='<%#Container.DataItem("MB_CLASS_NAME")%>' CommandArgument='<%#Container.DataItem("MB_SEQ")%>' Style="text-decoration:underline" CommandName="CONTENT" />                                                     
                                                </td>
                                            </tr>
                                        </ItemTemplate>
                                    </asp:Repeater>
                                    </asp:PlaceHolder>

                                    <tr>
                                        <td width="100%" class="NEWSTitle" colspan="2">
                                            <asp:LinkButton ID="LB_CLASS_2" runat="server" Text="課程活動預告" ForeColor="White" style="text-decoration:underline" />
                                        </td>                
                                    </tr>
                                    <asp:PlaceHolder ID="PLH_CLASS_2" runat="server" Visible="false">
                                    <asp:Repeater ID="RP_CLASS_2" runat="server">
                                        <ItemTemplate>
                                            <tr>
                                                <td width="50%" class="NEWSTd2c">
                                                    <!--課程名稱-->
                                                    <asp:LinkButton ID="LinkButton1" runat="server" Text='<%#Container.DataItem("MB_CLASS_NAME")%>' CommandArgument='<%#Container.DataItem("MB_SEQ")%>' Style="text-decoration:underline" CommandName="CONTENT" />                                                     
                                                </td>
                                            </tr>
                                        </ItemTemplate>
                                    </asp:Repeater>
                                    </asp:PlaceHolder>
                                </table>
                            </td>
                        </tr>
                    </table>
                    </asp:PlaceHolder>
                    
                    <asp:Repeater ID="RP_NEWS" runat="server">
                        <ItemTemplate>
                            <div style="padding-bottom:20px;float:left">
                            <table style="border-collapse:collapse;border-style:solid;empty-cells:show;border-width:1px;border-color:black;background:transparent;" cellspacing="0" cellpadding="0" align="center" width="1050px" >
                                <tr>
                                    <td width="50%" style="font-weight:bold;text-align:left;background:transparent;">
                                        <div style="font-size:16pt;font-weight:bold;display:none" >
                                            <asp:Literal ID="LTL_TITLE" runat="server" />
                                        </div>                            
                                        <asp:Literal ID="LTL_YYYYMMDD" runat="server" Text='<%#getTIME(Container.DataItem("CHGTIME"))%>' Visible="false" />
                                    </td>
                                    <td width="50%" style="text-align: right;background:transparent;">
                                        <input type="hidden" id="HID_CRETIME" runat="server" value='<%#Container.DataItem("CRETIME")%>' />
                                        <input type="hidden" id="HID_CHGTIME" runat="server" value='<%#Container.DataItem("CHGTIME")%>' />
                                        <input type="hidden" id="HID_SEQTIME" runat="server" value='<%#Container.DataItem("SEQTIME")%>' />

                                        <asp:PlaceHolder ID="PLH_ADMIN" runat="server" Visible=false>
                                            <asp:ImageButton ID="IMG_TOP" runat="server" CommandName="TOP" AlternateText="文章置頂" title="文章置頂" />
                                            <asp:ImageButton ID="IMG_UP" runat="server" CommandName="UP" AlternateText="往較新文章排序" title="往較新文章排序" />
                                            <asp:ImageButton ID="IMG_DOWN" runat="server" CommandName="DOWN" AlternateText="往較舊文章排序" title="往較舊文章排序" />
                                            <asp:ImageButton ID="IMG_ADD" runat="server" CommandName="ADD" AlternateText="新增文章" title="新增文章" />
                                            <asp:ImageButton ID="IMG_EDIT" runat="server" CommandArgument='<%#Container.DataItem("CRETIME")%>' CommandName="EDIT" AlternateText="修改" title="修改" />
                                            <asp:ImageButton ID="IMG_DEL" runat="server" onmousedown="if (!confirm('確定要刪除嗎?')){return false}else{this.click()}" CommandArgument='<%#Container.DataItem("CRETIME")%>' CommandName="DEL" AlternateText="刪除" title="刪除" />
                                            <input type="hidden" id="HID_CODEID" runat="server" value='<%#Container.DataItem("CODEID")%>' />
                                            <input type="hidden" id="HID_BRID" runat="server" />
                                        </asp:PlaceHolder>
                                    </td>
                                </tr>
                                <tr>
                                    <td style="text-align: left;background:transparent;" colspan=2 >
                                        <Table width="100%" border="0" cellspacing="0" cellpadding="0" style="width:1050px" >
                                            <tr>
                                                <td width="80%" style="vertical-align:top;background:transparent">
                                                    <asp:Literal ID="LTL_CNTHTML" runat="server" />
                                                </td>
                                                <td width="20%" style="vertical-align:top;background:transparent">
                                                    <Div id="DIV_MORE" runat="server" visible="false" style="text-align:right" >
                                                        <asp:Button ID="btnReadMore" runat="server" Text="READ MORE" CssClass="bt" Font-Bold="true" CommandName="MORE" />
                                                    </Div>
                                                    <asp:Panel ID="PNL_CLASS" runat="server" Visible="false">
                                                        <Table width="100%" border="0" cellspacing="2" cellpadding="0" style="text-align:right">
                                                            <Div style="text-align:right">
                                                                課程編號:<asp:Label ID="LTL_CLASS_MB_SEQ" runat="server" Text='<%#Container.DataItem("MB_SEQ")%>' />
                                                            </Div>                                                            
                                                            <td class="td" style="vertical-align:top;background:transparent">
                                                                <input type="hidden" id="HID_MB_SEQ" runat="server" value='<%#Container.DataItem("MB_SEQ")%>' />
                                                                <asp:Repeater ID="RP_SIGN" runat="server" OnItemDataBound="RP_SIGN_OnItemDataBound" OnItemCommand="RP_SIGN_OnItemCommand">
                                                                    <ItemTemplate>
                                                                        <DIV>
                                                                        <asp:Label ID="LTL_MB_BATCH" runat="server" Style="vertical-align:top" />
                                                                        <input type="hidden" id="HID_MB_SEQ" runat="server" value='<%#Container.DataItem("MB_SEQ")%>' />
                                                                        <input type="hidden" id="HID_MB_BATCH" runat="server" value='<%#Container.DataItem("MB_BATCH")%>' />
                                                                        <asp:PlaceHolder ID="PLH_CLASS_SIGN" runat="server" >
                                                                        <asp:Button ID="btnClass" runat="server" Text="我要報名" CommandName="CLASS" Font-Bold="true" CssClass="bt_sign"   />
                                                                        <asp:Button ID="btnCanClass" runat="server" Text="取消報名" CommandName="CANCLASS" Font-Bold="true" CssClass="bt_sign"   />
                                                                        </asp:PlaceHolder>
                                                                        <asp:PlaceHolder ID="PLH_APLY" runat="server">
                                                                            <asp:Label ID="LTL_APLY" runat="server" Style="font-size:16pt;font-weight:bold" />
                                                                        </asp:PlaceHolder>
                                                                        </DIV>
                                                                    </ItemTemplate>
                                                                </asp:Repeater>
                                                            </td>
                                                        </Table>                                           
                                                    </asp:Panel>
                                                </td>
                                            </tr>
                                        </Table>
                                    </td>
                                </tr>
                            </table>
                            </div>
                        </ItemTemplate>
                    </asp:Repeater>
                    <table border="0" cellspacing="0" cellpadding="0" align="center" width="960px">
                        <tr>
                            <td align="left" width="20%" style="background:transparent" >
                                <input type="hidden" id="HID_PAGE" runat="server" />
                                <asp:ImageButton ID="IMG_LEFT" runat="server" AlternateText="較新的文章" title="較新的文章" />
                            </td>
                            <td align="center" width="60%" style="background:transparent" >
                                <asp:ImageButton ID="IMG_ADD" runat="server" AlternateText="新增文章" Visible="false" title="新增文章" />
                            </td>                
                            <td align="right" width="20%" style="background:transparent" >
                                <asp:ImageButton ID="IMG_RIGHT" runat="server" AlternateText="較舊的文章" title="較舊的文章" />
                            </td>
                        </tr>
                    </table>
                    <hr style="margin: 5px 0!important; margin: 0px; height: 1px; overflow: hidden; border: #000000;background-color: #000000; color: #000000;" />
                    <asp:PlaceHolder ID="PLH_CKEditor" runat="server" Visible="false">
                        <div style='background:transparent;'>
                            <span style="font-weight: bold">YYYYMMDD:<asp:Literal ID="LTL_YYYYMMDD" runat="server" /></span><br />
                            <span style="font-weight: bold">HHMMSS:<asp:Literal ID="LTL_HHMMSS" runat="server" /></span>
                            <input type="hidden" id="HID_CRETIME" runat="server" />
                            <input type="hidden" id="HID_SEQTIME" runat="server" />
                            <input type="hidden" id="HID_INDEX" runat="server" />
                            <input type="hidden" id="HID_CODEID" runat="server" />
                            <div align="left">
                                選擇課程<asp:RadioButton ID="RB_CLASS_YES" runat="server" Text="是" GroupName="CLASS" onclick="document.all('btnChoose').style.display='';" />
                                <asp:RadioButton ID="RB_CLASS_NO" runat="server" Text="否" GroupName="CLASS" onclick="document.all('btnChoose').style.display='none';" />
                                <asp:Button ID="btnChoose" runat="server" CssClass="bt" Text="選擇課程" style="display:none" />
                            </div>
                            <h3 style="font-family:標楷體">簡介</h3>
                            <CKEditor:CKEditorControl ID="CKEditorControl2" runat="server" />
                            <h3 style="font-family:標楷體">內文</h3>
                            <CKEditor:CKEditorControl ID="CKEditorControl1" runat="server" />
                            <script language="javascript" type="text/javascript" >
                                // Replace the <textarea id="editor1"> with a CKEditor
                                // instance, using default configuration.
                                $(document).ready(function () {
                                    CKEDITOR.config.extraPlugins += (CKEDITOR.config.extraPlugins ? ',lineheight' : 'lineheight');

                                    CKEDITOR.replace('CKEditorControl1', {
                                        filebrowserBrowseUrl: getRootPath() + '/ckfinder/ckfinder.html', //不要寫成"~/ckfinder/..."或者"/ckfinder/..."
                                        filebrowserImageBrowseUrl: getRootPath() + '/ckfinder/ckfinder.html?Type=Images',
                                        filebrowserFlashBrowseUrl: getRootPath() + '/ckfinder/ckfinder.html?Type=Flash',
                                        filebrowserUploadUrl: getRootPath() + '/ckfinder/core/connector/aspx/connector.aspx?command=QuickUpload&type=Files',
                                        filebrowserImageUploadUrl: getRootPath() + '/ckfinder/core/connector/aspx/connector.aspx?command=QuickUpload&type=Images',
                                        filebrowserFlashUploadUrl: getRootPath() + '/ckfinder/core/connector/aspx/connector.aspx?command=QuickUpload&type=Flash',
                                        filebrowserWindowWidth: '800',  //“瀏覽服務器”彈出框的size設置
                                        filebrowserWindowHeight: '800',
                                        htmlEncodeOutput: true,
                                        height: '550px',
                                        skin: 'moono-dark',
                                        toolbarGroups: [
		                                                { name: 'clipboard', groups: ['clipboard', 'undo'] },
		                                                { name: 'editing', groups: ['find', 'selection', 'spellchecker'] },
		                                                { name: 'links' },
		                                                { name: 'insert' },
		                                                { name: 'forms' },
		                                                { name: 'tools' },
		                                                { name: 'document', groups: ['mode', 'document', 'doctools' ] },
		                                                { name: 'others' },
		                                                '/',
		                                                { name: 'basicstyles', groups: ['basicstyles', 'cleanup'] },
		                                                { name: 'paragraph', groups: ['list', 'indent', 'blocks', 'align', 'bidi' ] },
		                                                { name: 'styles' },
		                                                { name: 'colors' },
                                                        { name: 'lineheight' },
		                                                { name: 'about' }
                                                        ]
                                    });

                                    // Replace the <textarea id="editor1"> with a CKEditor
                                    // instance, using default configuration.
                                    CKEDITOR.replace('CKEditorControl2', {
                                        filebrowserBrowseUrl: getRootPath() + '/ckfinder/ckfinder.html', //不要寫成"~/ckfinder/..."或者"/ckfinder/..."
                                        filebrowserImageBrowseUrl: getRootPath() + '/ckfinder/ckfinder.html?Type=Images',
                                        filebrowserFlashBrowseUrl: getRootPath() + '/ckfinder/ckfinder.html?Type=Flash',
                                        filebrowserUploadUrl: getRootPath() + '/ckfinder/core/connector/aspx/connector.aspx?command=QuickUpload&type=Files',
                                        filebrowserImageUploadUrl: getRootPath() + '/ckfinder/core/connector/aspx/connector.aspx?command=QuickUpload&type=Images',
                                        filebrowserFlashUploadUrl: getRootPath() + '/ckfinder/core/connector/aspx/connector.aspx?command=QuickUpload&type=Flash',
                                        filebrowserWindowWidth: '800',  //“瀏覽服務器”彈出框的size設置
                                        filebrowserWindowHeight: '800',
                                        htmlEncodeOutput: true,
                                        height: '150px',
                                        skin: 'moono-dark',
                                        toolbarGroups: [
		                                                { name: 'clipboard', groups: ['clipboard', 'undo'] },
		                                                { name: 'editing', groups: ['find', 'selection', 'spellchecker'] },
		                                                { name: 'links' },
		                                                { name: 'insert' },
		                                                { name: 'forms' },
		                                                { name: 'tools' },
		                                                { name: 'document', groups: ['mode', 'document', 'doctools' ] },
		                                                { name: 'others' },
		                                                '/',
		                                                { name: 'basicstyles', groups: ['basicstyles', 'cleanup'] },
		                                                { name: 'paragraph', groups: ['list', 'indent', 'blocks', 'align', 'bidi' ] },
		                                                { name: 'styles' },
		                                                { name: 'colors' },
                                                        { name: 'lineheight' },
		                                                { name: 'about' }
                                                        ]
                                    });
                                })
                            </script>

                            <asp:Label ID="LBL_HTML" runat="server" />
                        </div>
                        <div align="center">
                            <asp:Button ID="btPreview" runat="server" Text="預覽" CssClass="bt" Visible="false" />
                            <asp:Button ID="btSAVE" runat="server" Text="儲存" CssClass="bt" />&nbsp;
                            <asp:Button ID="btCancel" runat="server" Text="取消" CssClass="bt" />
                        </div>
                    </asp:PlaceHolder>        
                </td>
            </tr>
        </table>
    </form>
    
    <script language="JavaScript" type="text/javascript" src="<%=com.Azion.EloanUtility.UIUtility.getRootPath() + "/js/jquery.bxslider.min.js"%>"></script>

    <script language="javascript" type="text/javascript">
        //if (window.chrome) {
        //    $('.banner li').css('background-size', '100% 100%');
        //}

        //$('.banner').unslider({
        //    speed: 500,               //  The speed to animate each slide (in milliseconds)
        //    delay: 3000,              //  The delay between slide animations (in milliseconds)
        //    complete: function () { },  //  A function that gets called after every slide animation
        //    keys: true,               //  Enable keyboard (left, right) arrow shortcuts
        //    dots: true,               //  Display dot navigation
        //    fluid: false              //  Support responsive design. May break non-responsive designs
        //});

        $(document).ready(function () {
            $('.bxslider').bxSlider({
                auto: true,
                captions: true,
                controls:false
            });
        });
    </script>
</body>
</html>
