<html>
<head>
<title>�˳�ϵͳ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body bgcolor="#EFEFEF">
<boday>
  <div align=center>
    <center>
      <table width="100%" height="100%" vAlign=middle align=center>
        <tr>
          <td align=center><span id=Logout style="font-size:10pt;color:#FF0000;font-family:verdana;"><font color="#FF6600">�ɹ��˳�ϵͳ</font></span>
            <script language="JavaScript">
	var n=5;
	function Logout() {
		document.all("Logout").innerText = "�ɹ��˳�ϵͳ���Ժ��Զ���ת����ҳ(" + n + ")";
		if( n == 2 || n == 1 )
			document.all("Logout").innerText = "����׼����ת(" + n + ")";
		if( n == 0 )
			document.all("Logout").innerText = "������ҳ����";

		
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
