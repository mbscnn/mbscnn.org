<style>
#nav {
	width:1200px;
	height:36px;
	line-height:36px;
	/*background:#E2D4B0;*/
  /*background:url(<%=com.Azion.EloanUtility.UIUtility.getRootPath() + "/img/mbscbg3.gif"%>) repeat;*/
  background:url(<%=com.Azion.EloanUtility.UIUtility.getRootPath() + "/img/PageTabBg.gif"%>) no-repeat;
  /*background:transparent;*/
  margin-left:auto;
  margin-right:auto;
  position:relative;
  z-index: 1000;
  border:1px solid black;
}     

#nav ul, #nav ul li {
  margin:0;
  padding:0;
  list-style:none;
}

#nav ul li {
  float:left;
  display:block;
}

#nav ul li a:link, #nav ul li a:visited 
{
	color:white;
	font-size:12px;
	font-weight:bold;
	text-decoration:none;
	display:block;
	text-align: center;
	width:87px;
	height:36px;
	overflow:hidden;
	/*background:url(<%=com.Azion.EloanUtility.UIUtility.getRootPath() + "/img/mbscbg.gif"%>) no-repeat;*/
}

#nav ul li a:hover 
{
	/*color:#EBEFF7;*/
	/*background:url(<%=com.Azion.EloanUtility.UIUtility.getRootPath() + "/img/mbscbg.gif"%>) no-repeat;*/
  color: #de8529;
}

#nav ul li ul li 
{
  float:none;
  display:block;
}

#nav ul li ul li a:link, #nav ul li ul li a:visited 
{
  width:157px;
  font-size:12px;
  font-weight:bold;
  text-decoration:none;
  background:black;
  clear:both;
}

#nav ul li ul li a:hover
{
  /*color:#3B5998;*/
  width:157px;
  /*background:url(<%=com.Azion.EloanUtility.UIUtility.getRootPath() + "/img/mbscbg.gif"%>) no-repeat;*/
	color: #de8529;
}

.submenu 
{
	position: absolute;
	width: 157px;
	border: none;
	display: none;
	line-height: 36px;
	z-index: 1001;
}

.status-msg-wrap {
		width:1050px;
	  margin-left:auto;
	  margin-right:auto;		
    font-size: 110%;
    /*margin: 10px auto;*/
    position: relative;
    z-index: 999;
    clear:both;
}

.status-msg-body {
    padding: .3em 0;
    width: 1050px;
    position: absolute;
    z-index: 4;
    font-weight:bold;
    margin:0 auto;
    text-align:center;
}

.status-msg-border {
    border: 1px solid #000;
    filter: alpha(opacity=40);
    -moz-opacity: .4;
    opacity: 0.4;
    width: 1050px;
    position: relative;
}

.status-msg-bg {
    background-color: #ccc;
    opacity: 0.8;
    filter: alpha(opacity=30);
    -moz-opacity: .8;
    width: 1050px;
    position: relative;
    z-index: 1;
}

.status-msg-hidden {
    visibility: hidden;
    padding: .3em 0;
}
</style>

<!--<script type="text/javascript">
    function nav() {
        $('div#nav ul li').mouseover(function () {
            $(this).find('ul:first').show();
        });

        $('div#nav ul li').mouseleave(function () {
            $('div#nav ul li ul').hide();
        });

        $('div#nav ul li ul').mouseleave(function () {
            $('div#nav ul li ul').hide(); ;
        });
    };

    $(document).ready(function () {
        nav();
    });
</script>-->

<%@ Register src="~/Module/PageTab/PageTab.ascx" tagname="PageTab" tagprefix="uc1" %>
    
<uc1:PageTab ID="PageTab1" runat="server" />