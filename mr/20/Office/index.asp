<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!-- #include file="Conn/conn.asp" -->  
<!--�������ݿ������ļ�-->
<% 
    session.Timeout=120 
if request.Form("UserName")<>"" and request.Form("PWD")<>"" then
	session("UserName")=request.Form("UserName")
	session("PWD")=request.Form("PWD")
	sql="select UserName,PWD from Tab_User where UserName='"&session("UserName")&"'"
	set rs=conn.execute(sql)
	if rs.eof then 
%>
  		<script language="javascript">
  		alert("��������û����������������룡");
 		 </script> 
 		 <%
		 session.Abandon()  
		 'ɾ�����д���Session�����еĶ���
	else 
		if rs("PWD")=session("PWD") then
			session("flag")="��¼" %>
     		<script language="javascript">
      		window.location.href="default.asp"
  	 		</script>
 		 <%else%>
   			<script language="javascript">
			 alert("��������û�����������������룡");
  			</script>  		
			<%session.Abandon()
		end if
	end if
end if
%>
<%   
    '�Զ�����4λ�������֤��
	randomize
	i=0
	do while i<4
	num1=int(9*rnd)
	numimage="<img src=images/num/"&num1&".gif>"
	numi=numi&numimage
	num=num&num1
	i=i+1
	loop
	session("numi")=numi
	session("num")=num
%>
<script language="javascript">
function mycheck(){
if (form1.UserName.value=="")
{alert("�������û�����");form1.UserName.focus();return;}
if(form1.PWD.value=="")
{alert("���������룡");form1.PWD.focus();return;}
if(form1.yanzheng.value=="")
{alert("��������֤��!");form1.yanzheng.focus();return;}
if(form1.yanzheng.value!=form1.verifycode2.value)
{alert("��������ȷ����֤��!!");form1.yanzheng.focus();return;}
form1.submit();
}
</script>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>��ҵ�칫�Զ���ϵͳ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="style.css" rel="stylesheet">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	background-color: #24A7F5;
}
.style2 {color: #990000}
.input2 {
	font-size: 12px;
	border: 3px double #A8D0EE;
	color: #344898;
}
.submit1 {
	border: 3px double #416C9C;
	height: 22px;
	width: 45px;
	background-color: #F2F2F2;
	font-size: 12px;
	padding-top: 1px;
	background-image: url(bt.gif);
	cursor: hand;
}
.STYLE12 {font-family: Georgia, "Times New Roman", Times, serif}
.STYLE13 {color: #316BD6; }
.STYLE15 {color: #316BD6; font-size: 9pt; }
-->
</style></head>

<body><br><br><br><br>
<table width="699" height="423" border="0" align="center" cellpadding="0" cellspacing="0" background="Images/denglu.jpg">
  <tr>
    <td width="706">
	  <form name="form1" method="POST" action="index.asp">
	    <table width="252" height="244" border="0" align="center" cellpadding="0" cellspacing="0">
	      <tr>
	        <td height="2" colspan="2"></td>
          </tr>
	      <tr>
	        <td height="97" colspan="2" valign="bottom"><div align="center" class="STYLE15">==== ϵͳ��¼ ====</div></td>
          </tr>
	      <tr>
	        <td width="54" height="22" valign="bottom"><span class="STYLE15">�û�����</span></td>
            <td width="200" valign="bottom"><input name="UserName" type="text" class="input2" id="userName" 
			onKeyDown="if(event.keyCode==13){form1.PWD.focus();}"onMouseOver="this.style.background='#F0DAF3';" onMouseOut="this.style.background='#FFFFFF'"></td>
          </tr>
	      <tr>
	        <td height="4" colspan="2" valign="bottom"></td>
          </tr>
	      <tr>
	        <td height="31" colspan="2" valign="top" class="STYLE15">��&nbsp;&nbsp;�룺            
            <input name="PWD" type="password" class="input2" id="PWD" align="bottom"
			 onKeyDown="if(event.keyCode==13){form1.yanzheng.focus();}"onMouseOver="this.style.background='#F0DAF3';" onMouseOut="this.style.background='#FFFFFF'"></td>
          </tr>
	      <tr>
	        <td height="31" colspan="2" valign="top" class="STYLE15"  ondragstart="return false" onselectstart="return false">��֤�룺            
	          <input name="yanzheng" type="text" class="input2" id="yanzheng"
			 onKeyDown="if(event.keyCode==13){form1.Submit.focus();}" size="8" align="bottom"onMouseOver="this.style.background='#F0DAF3';" onMouseOut="this.style.background='#FFFFFF'">
	          <input type="hidden" name="verifycode2" value="<%=session("num")%>">
            <span class="STYLE12"><font size="+3" color="#FF0000"><%=session("numi")%></font></span></td>
          </tr>
	      <tr>
	        <td colspan="2" valign="top">
	          &nbsp; &nbsp; &nbsp; &nbsp; <input name="Submit" type="button" class="submit1" value="��¼" onClick="mycheck()">
  &nbsp;   <input name="Submit2" type="reset" class="submit1" value="����">          </td>
          </tr>
        </table>
      </form>
	</td>
  </tr>
</table>
</body>
</html>
