<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Conn/conn.asp"-->
<!--#include file="../manage/include/md5.asp" -->
<%
flag = true			'û�д������true���д������false
str = ""			'������Ϣ'
jiaoyanma = Trim(Request.form("jiaoyanma"))
kaohao = Trim(request.form("kaohao"))
act = Trim(request.Form("act"))
If jiaoyanma <> "" Then
	set rs_1=server.CreateObject("adodb.recordset")
	Sql_1 = "Select * from tb_jiaoyanma Where jiaoyanma = '"&jiaoyanma&"'"
	Set rs_1 = conn.Execute(Sql_1)
	if rs_1.Eof then
		flag = false
		str = "�Բ���У���벻��ȷ!"
	end if 
	session("jiaoyanma") = jiaoyanma
	rs_1.close()
end if
if	kaohao <> "" Then
	set rs_2=server.CreateObject("adodb.recordset")
	Sql_2 = "select * from tb_StuResult where ����='"&kaohao&"'"
	Set rs_2 = conn.Execute(Sql_2)
	if rs_2.Eof then
		flag = false
		str = "�ÿ����ɼ���û�й��������߸ÿ��Ų�����!"
	else
		id=rs_2("id")
		dim rndnum,verifycode
		Randomize
		Do While Len(rndnum)<8
			num1=CStr(Chr((57-48)*rnd+48))
			rndnum=rndnum&num1
		loop
	
		'������¼��������ȡ����������ӵ����ݱ�tb_Student��
		set rs1=server.CreateObject("adodb.recordset")
		Sql1 = "select * from tb_Student "
		rs1.Open Sql1,conn,1,3
		rs1.AddNew
		rs1("UserID")=kaohao
		rs1("PWD") = rndnum
		rs1("PWD2") = md5(trim(rndnum))
		rs1("Jointime")=now()
		rs1.Update
		rs_2.close()
		'��ʹ�ú��У����ɾ��
		if session("jiaoyanma") <> "" then
			Sql_3 = "delete from tb_jiaoyanma Where jiaoyanma = '"&session("jiaoyanma")&"'"
			Set rs_3 = conn.Execute(Sql_3)
		end if
		'rs_3.close()
	end if
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��ȡ����</title>
<link  rel="stylesheet" type="text/css" href="Css/style.css">
<style type="text/css">
td{
	text-align:center;
	line-height:25px;
}
</style>
</head>
<script src="Js/Check.js"></script>
<script type="text/javascript">
//document.onkeydown = enterkey;
function enterkey(){
	if(event.keyCode == 116){				//��������ǡ�F5����
		alert('��ֹˢ��');					//���������
		event.keyCode = 0;					//����������
		return false;
	}
}
</script>
<body topmargin="0" leftmargin="0"  oncontextmenu=self.event.returnValue=false>
<div style=" height: 5px;">&nbsp;</div>
<% if flag = false then %>
<table border="0" cellpadding="0" cellspacing="0" align="center">
	<tr>
		<td>��������</td>
	</tr>
	<tr>
		<td>
		<%
				response.Write(str)
			
		%>
		</td>
	</tr>
	<tr>
		<td><a style="cursor:hand;" onClick="history.go(-1);">[����]</a></td>
	</tr>
</table>
<% else %>
<table border="0" cellpadding="0" cellspacing="0" align="center">
<form name="addPass_1" method="post" action="Default.asp" onSubmit="return Check(addPass_1)">
<%
	if act = "" then
%>
	<tr>
		<td>������У����</td>
	</tr>
	
	<tr>
		<td>�����룺<input name="jiaoyanma" type="text" class="text"></td>
	</tr>
	<tr>
		<td><input name="act" value="1" type="hidden" /><input name="�ύ" type="submit" class="btn_grey" value="��һ��"></td>
	</tr>
<%
	elseif act = "1" then
%>
	<tr>
		<td>������׼��֤��</td>
	</tr>
	
	<tr>
		<td>�����룺<input name="kaohao" type="text" class="text"></td>
	</tr>
	<tr>
		<td><input name="act" value="2" type="hidden" /><input name="�ύ" type="submit" class="btn_grey" value="��һ��"></td>
	</tr>
<%	elseif act = "2" then
%>
	<tr>
		<td>ע��ɹ�</td>
	</tr>
	
	<tr>
		<td>���������ǣ�<input name="mima" type="text" class="text" readonly="readonly" style="background-color: #FFFFCC; border: 1px #f0f0f0 dashed; color:#cc0000" value="<%=rndnum%>"></td>
	</tr>
	<tr>
		<td>����ҳ������ʹ��׼��֤�ź͸������¼����ѯ,ף��ȡ�úóɼ���</td>
	</tr>
<%
	else
	end if
%>
</form>
</table>
<% end if %>
</body>
</html>
