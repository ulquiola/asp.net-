<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
</head>
<body>
<%
user_login= Request.ServerVariables("URL")
If Request.QueryString<>"" Then user_login = user_login+"?"+Request.QueryString
username_value=CStr(Request.Form("UserName"))
application("UserName")=""
a_username=split(application("UserName"),",")
for i=0 to ubound(a_username)
	if a_username(i)=username_value then%>
	<script language="javascript">
	alert("�Բ��𣬸��û����ڱ𴦵�¼��");
	window.location.href="index1.asp";
	</script>
<%
response.End()
end if
next

If username_value <> "" Then
  user_action=""
  user_loginSuccessed="mainmenu.asp"
  user_loginfailed="index1.asp"
  set rsuser=Server.CreateObject("ADODB.Recordset")						'������¼��
  rsuser.ActiveConnection = conn										
  rsuser.Source = "SELECT Username, Password1,state,id,touxiang"		'��ѯ����
  If user_action <> "" Then rsuser.Source = rsuser.Source & "," & user_action
  rsuser.Source = rsuser.Source & " FROM tb_user where Username='" & Replace(username_value,"'","''") &"' and Password1='" & Replace(Request.Form("Password1"),"'","''")&"'"
  '�ж��������Ϣ�Ƿ���ȷ
  rsuser.CursorType = 0					'�����α�����
  rsuser.CursorLocation = 2				'�����α��������
  rsuser.LockType = 3					
  rsuser.Open							'�򿪼�¼��
  If Not rsuser.eof or not rsuser.bof Then 	'��ѯ�Ƿ����û���Ϣ
    Session("username1")=username_value		'ΪSession("username1")������ֵ
	application.Lock()						
	application("username")=application("username")+","+username_value	'Ϊapplication("username")������ֵ
	application.UnLock()
	response.Cookies("user_name") = username_value						'��ȡ��ǰ�û���
	response.Cookies("state") = rsuser.fields.item("state").value		'��ȡָ�����ֶ���Ϣ
	response.Cookies("UserID") = rsuser.fields.item("id").value			'��ȡָ�����ֶ���Ϣ
	Session("touxiang")=rsuser.fields.item("touxiang").value			'��ȡָ�����ֶ���Ϣ
	Session("user_id")=rsuser.fields.item("id").value					'��ȡָ�����ֶ���Ϣ
	Session("state")=rsuser.fields.item("state").value					'��ȡָ�����ֶ���Ϣ
	if trim(application("OnLineUserID")) = "" then 
		application("OnLineUserID") = "S" & rsuser.fields.item("id").value & "E"
	else
	    strUserID =  ",S" & rsuser.fields.item("id").value & "E"
	    application("OnLineUserID") = replace(application("OnLineUserID"),strUserID,"")
	    application("OnLineUserID") = application("OnLineUserID") & ",S" & rsuser.fields.item("id").value&"E"
	End IF
    If (user_action <> "") Then
      Session("user_action") = CStr(rsuser.Fields.Item(user_action).Value)
    Else
      Session("user_action")=""											'��� Session("user_action")����
    End If
    if CStr(Request.QueryString("accessdenied")) <> "" And false Then	
      user_loginSuccessed = Request.QueryString("accessdenied")		'Ϊuser_loginSuccessed������ֵ
    End If
    rsuser.Close
    Response.Redirect(user_loginSuccessed)					'��ת��ָ����ҳ��
  End If
  rsuser.Close
  Response.Redirect(user_loginfailed)						'��ת��ָ����ҳ��
End If
%>
</body>
</html>
