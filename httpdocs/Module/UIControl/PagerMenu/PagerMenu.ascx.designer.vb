'------------------------------------------------------------------------------
' <自動產生的>
'     這段程式碼是由工具產生的。
'
'     變更這個檔案可能會導致不正確的行為，而且如果已重新產生
'     程式碼，則會遺失變更。
' </自動產生的>
'------------------------------------------------------------------------------

Option Strict On
Option Explicit On


Partial Public Class PagerMenu
    
    '''<summary>
    '''tbPagingNavigation 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents tbPagingNavigation As Global.System.Web.UI.HtmlControls.HtmlTable
    
    '''<summary>
    '''lblPageIndex 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents lblPageIndex As Global.System.Web.UI.WebControls.Label
    
    '''<summary>
    '''span1 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents span1 As Global.System.Web.UI.HtmlControls.HtmlGenericControl
    
    '''<summary>
    '''span2 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents span2 As Global.System.Web.UI.HtmlControls.HtmlGenericControl
    
    '''<summary>
    '''lblTotalPage 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents lblTotalPage As Global.System.Web.UI.WebControls.Label
    
    '''<summary>
    '''lblTotalRecord 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents lblTotalRecord As Global.System.Web.UI.WebControls.Label
    
    '''<summary>
    '''imgBtnFirst 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents imgBtnFirst As Global.System.Web.UI.WebControls.ImageButton
    
    '''<summary>
    '''btnNavFirst 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents btnNavFirst As Global.System.Web.UI.WebControls.LinkButton
    
    '''<summary>
    '''spanFirst 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents spanFirst As Global.System.Web.UI.HtmlControls.HtmlGenericControl
    
    '''<summary>
    '''imgBtnPrevious 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents imgBtnPrevious As Global.System.Web.UI.WebControls.ImageButton
    
    '''<summary>
    '''btnNavPrevious 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents btnNavPrevious As Global.System.Web.UI.WebControls.LinkButton
    
    '''<summary>
    '''spanSecond 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents spanSecond As Global.System.Web.UI.HtmlControls.HtmlGenericControl
    
    '''<summary>
    '''btnNavNext 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents btnNavNext As Global.System.Web.UI.WebControls.LinkButton
    
    '''<summary>
    '''spanThird 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents spanThird As Global.System.Web.UI.HtmlControls.HtmlGenericControl
    
    '''<summary>
    '''imgBtnNext 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents imgBtnNext As Global.System.Web.UI.WebControls.ImageButton
    
    '''<summary>
    '''btnNavLast 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents btnNavLast As Global.System.Web.UI.WebControls.LinkButton
    
    '''<summary>
    '''imgBtnLast 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents imgBtnLast As Global.System.Web.UI.WebControls.ImageButton
    
    '''<summary>
    '''ddlJumpPage 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents ddlJumpPage As Global.System.Web.UI.WebControls.DropDownList
End Class
