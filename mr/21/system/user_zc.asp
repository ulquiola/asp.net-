<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn/conn.asp"-->
<%
if request.Form("usename")<>"" then
	usename=request.Form("usename")
	pwd=request.Form("pwd")
	purview=request.Form("purview")
	'���������û��Ƿ��Ѵ���
	set rs=server.CreateObject("Adodb.recordset")
	sql="select * from tb_user where usename='"&usename&"'"
	rs.open sql,conn,1,3
	if rs.eof then 
	'ͨ������ĳ�������������û�ע����Ϣ
	ins="insert into tb_user (usename,pwd,purview) values('"&usename&"','"&pwd&"','"&purview&"')"
	conn.execute(ins)
%>
	<script language="javascript">
	alert("�û���Ϣ��ӳɹ���������ӣ�");
	</script> 
	<%else%>
		<script language="javascript">
	    alert("���û���Ϣ�Ѿ����ڣ�");
	    </script> 
<%		end if
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�û�ע��</title>
<link rel="stylesheet" href="Css/style.css">
<style type="text/css">
<!--
.STYLE1 {color: #107483}
.STYLE2 {color: #595959}
body {
	margin-top: 0px;
	margin-bottom: 0px;
	background-image: url(images/bg.gif);
}
-->
</style>
</head>
<script language="javascript">
function Mycheck(form1)
{
  for(i=0;i<form1.length;i++)
   {
    if(form1.elements[i].value=="")
	{alert(form1.elements[i].title + "����Ϊ��!");return false;}		
    }
	form1.submit();
}
</script>

<body>
<p>&nbsp;</p>
<p>&nbsp;</p>
<table width="773" height="453" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="666" align="center" valign="top" background="images/zhece0.gif"><meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="Css/style.css" rel="stylesheet" />
<p>
  <script src="JS/onclock.JS"></script>
  <script src="JS/menu.JS"></script>
</p>
<table width="90%" height="270" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td width="5" valign="top">&nbsp;</td>
        <td align="right" valign="middle">
		<form method="post" name="form1">
		  <p>&nbsp;</p>
		  <p>&nbsp;</p>
		  <table width="426" height="153" border="0" align="right" cellpadding="0" cellspacing="0">
          <tr>
            <td width="196" align="center"><div align="center"><span class="STYLE2">�û����ƣ�</span></div></td>
            <td width="230"><input name="usename" type="text" id="usename" title="�û�����" size="25"></td>
          </tr>
          <tr>
            <td align="center" class="STYLE1"><div align="center" class="STYLE2">�û����룺</div></td>
            <td><input name="pwd" type="password" id="pwd" title="�û�����" size="25"></td>
          </tr>
          <tr>
            <td colspan="2" align="center" class="STYLE1">
			<input type="hidden" name="purview" class="STYLE2" id="purview" value="��ͨ�û�">
            </td>
            </tr>
          <tr>
            <td height="66" colspan="2"><p> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="images/button1.gif" width="56" height="25" onClick="Mycheck(form1)">                <img src="images/button2.gif" width="56" height="25" onClick="Jscript:form1.reset();"><img src="images/button3.gif" width="59" height="25" onClick="window.location.href='index.asp';"></p>              </td>
            </tr>
        </table>
		</form>		</td>
        </tr>
    </table>
<style type="text/css">
<!--
.STYLE1 {color: #FFFFFF}
-->
</style>
<div class=menuskin id=popmenu
      onmouseover="clearhidemenu();highlightmenu(event,'on')"
      onmouseout="highlightmenu(event,'off');dynamichide(event)" style="Z-index:100;position:absolute;"></div>    </td>
  </tr>
</table>
</body>
</html>
