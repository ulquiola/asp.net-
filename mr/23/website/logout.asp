<html>
<head>
<title>退出系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body bgcolor="#EFEFEF">
<boday>
  <div align=center>
    <center>
      <table width="100%" height="100%" vAlign=middle align=center>
        <tr>
          <td align=center><span id=Logout style="font-size:10pt;color:#FF0000;font-family:verdana;"><font color="#FF6600">成功退出系统</font></span>
            <script language="JavaScript">
	var n=5;
	function Logout() {
		document.all("Logout").innerText = "成功退出系统，稍后将自动跳转到首页(" + n + ")";
		if( n == 2 || n == 1 )
			document.all("Logout").innerText = "正在准备跳转(" + n + ")";
		if( n == 0 )
			document.all("Logout").innerText = "返回首页……";

		
		if( !n-- )
			this.location = "index.asp";
		setTimeout( "Logout();", 1000 );
	}
	Logout();
</script>
          </td>
        </tr>
      </table>
    </center>
  </div>
</boday>
</html>
