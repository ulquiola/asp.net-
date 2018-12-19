<% 
  Response.Expires = -1  
  Response.ExpiresAbsolute = Now() - 1  
  Response.cachecontrol = "no-cache"  
  Session.TimeOut = 200
%>
<html>
<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<title>系统</title>
</head>
<frameset frame rows="128,*,25" cols="*" framespacing="1" frameborder="Yes" border="1">
  <frame name="EwSys_Top" scrolling="no" noresize target="contents" src="admin_top.asp?EwSysShow=EwSys_Top">
  <frameset id=EwSysTem cols="180,*" framespacing="0" frameborder="no" border="0">
    <frame name="EwSys_Left" target="main" src="Admin_Left.asp" scrolling="auto" marginwidth="0" marginheight="0">
    <frame name="EwSys_Main" scrolling="auto" noresize src="server.asp">
  </frameset>
  <noframes>
  <body>
  <p>此网页使用了框架，但您的浏览器不支持框架.</p>
  </body>
  </noframes>
</html>
