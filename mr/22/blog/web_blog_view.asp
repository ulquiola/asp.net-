<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="include/conn.asp"-->
<!--#include file="include/function.asp"-->
<%
Session("verify")=""
'����¼�¼
Sub add()
  id=Request.Form("id")
  str1=Str_filter(Request.Form("�������ǳ�")) 
  str2=Str_filter(Request.Form("��������")) 
If str1<>"" and str2<>"" Then
  Set rs=Server.CreateObject("ADODB.Recordset")
  sqlstr="select * from tab_article_commend"
  rs.open sqlstr,conn,1,3
  rs.addnew
  rs("Cid")=id
  rs("Cname")=str1
  rs("Ccontent")=str2
  rs.update
  rs.close
  Set rs=Nothing
  Response.Redirect("web_blog_view.asp?id="&id&"")
Else
  Response.Write("<script>alert('����д����Ϣ������!');history.back();</script>")
End IF
End Sub
If Not Isempty(Request("add")) Then call add()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>����</title>
<link rel="stylesheet" href="css/css.css" />
<script language="javascript">
function Mycheck(form){
  for(i=0;i<form.length;i++){
    if(form.elements[i].value==""){
	  alert(form.elements[i].name + "����Ϊ��!");return false;}	
	if(form.elements[1].value!=form.elements[2].value){
	  alert("��֤�����!");return false;}  
  }
}
</script>
<SCRIPT language=JavaScript>
<!--
var LastCount =0;
function CountStrByte(Message,Total,Used,Remain){ //�ֽ�ͳ��
 var ByteCount = 0;
 var StrValue  = Message.value;
 var StrLength = Message.value.length;
 var MaxValue  = Total.value;

 if(LastCount != StrLength) { // �ڴ��жϣ�����ѭ������
	for (i=0;i<StrLength;i++){
		ByteCount  = (StrValue.charCodeAt(i)<=256) ? ByteCount + 1 : ByteCount + 2;
      if (ByteCount>MaxValue) {
      	Message.value = StrValue.substring(0,i);
			alert("����������಻�ܳ��� " +MaxValue+ " ���ֽڣ�\nע�⣺һ������Ϊ���ֽڡ�");
         ByteCount = MaxValue;
         break;
      }
	}
   Used.value = ByteCount;
   Remain.value = MaxValue - ByteCount;
   LastCount = StrLength;
 }
}
//-->
</SCRIPT>
</head>
<body>
<table width="900" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td colspan="2"><!--#include file="top.asp"--></td>
  </tr>
  <tr>
    <td width="210" valign="top" bgcolor="#FFFFFF"><!--#include file="left.asp"--></td>
    <td width="565" valign="top" bgcolor="#FFFFFF">
      <table width="100%"  border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td bgcolor="#FFFFFF"><!--#include file="Item.asp"-->
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
<%
'��ʾ���µ���ϸ����,�������±��⡢�������ߡ�����ʱ���Լ���������
id=Request.QueryString("id")
classname=Request.QueryString("classname")
If Not IsNumeric(id) Then
Else
Set rs=conn.Execute("select id,Atitle,Acontent,Aauthor,Adate from tab_article where id="&id&"")
%>
  <tr align="left" bgcolor="FFFCE8">
    <td colspan="2"><font color="#FF0000"><%=classname%></font> -&gt; <%=rs("Atitle")%></td>
  </tr>
  <tr>
    <td colspan="2" align="center" class="font1"><%=rs("Atitle")%></td>
  </tr>
  <tr>
    <td width="58%" align="right">���ߣ�<%=rs("Aauthor")%> </td>
    <td width="42%" align="center"><%=rs("Adate")%></td>
  </tr>
  <tr>
    <td colspan="2" style="line-height:22px">&nbsp;&nbsp;<%=rs("Acontent")%></td>
  </tr>
  <tr>
    <td colspan="2">
	<table width="100%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#ECECEC">
<%
'��ʾ�Դ�ƪ���·������ϸ��������
Set rsc=Server.CreateObject("ADODB.Recordset")
sqlstr="select top 25 * from tab_article_commend where Cid="&id&" order by id desc"
rsc.open sqlstr,conn,1,1
while not rsc.eof
%>
<form name="form<%=rsc("id")%>" method="post" action="">
	<tr bgcolor="#FFFFFF">
	  <td height="22"><font class="font1">�������ǳ�:</font><%=rsc("Cname")%></td>
	  <td height="22"><font class="font1">����ʱ��:</font><%=rsc("Cdate")%></td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td height="22" colspan="2"><font class="font1">��������:</font><%=rsc("Ccontent")%></td>
	  </tr>
</form>
<%
rsc.movenext
wend
rsc.close
Set rsc=Nothing
%>
  </table>
	</td>
  </tr>
  <tr>
    <td colspan="2"><table width="100%" border="0" cellspacing="0" cellpadding="0">
	<form name="form1" method="post" action="">
      <tr>
        <td width="23%" height="22" align="right" class="font2">�������ǳƣ�</td>
        <td width="77%" height="22"><input name="�������ǳ�" type="text" class="textbox" id="�������ǳ�" /></td>
      </tr>
      <tr>
        <td height="22" align="right" class="font2">��֤���룺</td>
        <td height="22">
		<%Session("verify")=randStr(4)%>
		<input name="��֤����" type="text" class="textbox" id="��֤����" size="6" />
		<font color="#FF0000"><%=Session("verify")%></font>
        <input name="verify2" type="hidden" id="verify2" value="<%=Session("verify")%>" /></td>
      </tr>
      <tr>
        <td height="22" align="right" class="font2">�������ݣ�</td>
        <td height="22"><textarea name="��������" cols="45" rows="5" id="��������" onKeyDown="CountStrByte(this.form.��������,this.form.total,this.form.used,this.form.remain);" onKeyUp="CountStrByte(this.form.��������,this.form.total,this.form.used,this.form.remain);"></textarea>
		  <br />
		  ������� <input name="total" type="text" disabled class="textbox" id="total"  value="200" size="3"> ���ֽ� �����ֽڣ�<input name="used" type="text" disabled class="textbox" id="used"  value="0" size="3">  ʣ���ֽڣ�<input name="remain" type="text" disabled class="textbox" id="remain" value="200" size="3">
		</td>
      </tr>
      <tr>
        <td height="22">&nbsp;</td>
        <td height="22"><input name="add" type="submit" class="button" id="add" onClick="return Mycheck(this.form)" value="�� ��" />
          <input name="Submit2" type="reset" class="button" value="�� ��" />
          <input name="id" type="hidden" id="id" value="<%=id%>" /></td>
      </tr>
	</form>
    </table></td>
  </tr>
 <%
Set rs=Nothing
End IF%> 
</table></td>
  </tr>
</table></td>
  </tr>
  <tr>
    <td colspan="2"><!--#include file="bottom.asp"--></td>
  </tr>
</table>
</body>
</html>
