<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!-- #include file="Conn/Conn.asp" -->
<!--�������ݿ������ļ�-->
<% 
Session.Timeout=30
function filter_Str(InString)
'***********************************
'���ܣ����������ַ����е�Σ�շ���
'���÷�����filter_Str("String")
'***********************************
	NewStr=Replace(InString,"'","''")
	NewStr=Replace(NewStr,"<","&lt;")
	NewStr=Replace(NewStr,">","&gt;")
	NewStr=Replace(NewStr,"chr(60)","&lt;")
	NewStr=Replace(NewStr,"chr(37)","&gt;")
	NewStr=Replace(NewStr,"""","&quot;")
	NewStr=Replace(NewStr,";",";;")
	NewStr=Replace(NewStr,"--","-")
	NewStr=Replace(NewStr,"/*"," ")
	NewStr=Replace(NewStr,"%"," ")
	filter_Str=NewStr
end function
if Request.Form("UID")<>"" and request.Form("PWD")<>"" then
	UID=filter_Str(request.Form("UID"))
	PWD=filter_Str(request.Form("PWD"))
	sql="select UserName,PWD from tb_User where UserName='"&UID&"'"
	set rs=conn.execute(sql)
	if rs.eof then
	%>
  		<script language="javascript">
  		alert("��������û����������������룡");
		window.location.href="Login.asp";
 		 </script> 
	<%
	Session("HS")="User"
	else 
		if rs("PWD")=PWD then
			Session("UserName")=UID
	lastlogintime=now()
	UP="Update tb_user set logintimes=logintimes+1,lastlogintime='"&lastlogintime&"' where UserName='"&UID&"'"
	conn.execute(UP)
	%>
	<script language=javascript>
 alert("��¼�ɹ���");
parent.leftwindow.location.reload();
window.location.href="main.asp";
</script>
	<%
 		 else
	%>
   			<script language="javascript">
			 alert("��������û�����������������룡");
			 window.location.href="Login.asp";
  			</script>  		
		<%
		Session("HS")="User"
		end if
	end if
end if
%>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

