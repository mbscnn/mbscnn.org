<%@ Page %>
<html>
<head>
<title>BOT e-loan sub-system main page</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">

</head>
  <frameset cols="160,*" frameborder="YES" border="1" rows="*">
    <frame name="leftFrame" src="<%=com.Azion.EloanUtility.UIUtility.getRootPath()%>/SY/SY_FUNCTIONLIST.aspx" scrolling="yes" style="border-right-style: groove;border-right-width: thin">
    <frame name="mainFrame" src="<%=com.Azion.EloanUtility.UIUtility.getRootPath()%>/SY/SY_CASELIST.aspx?ftype=1" scrolling="yes">
  </frameset>
<noframes> 
<body bgcolor="#FFFFFF" text="#000000">
</body>
</noframes> 
</html>
