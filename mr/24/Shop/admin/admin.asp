<%Option Explicit 
Dim ObjTotest(26,4)
ObjTotest(0,0) = "MSWC.AdRotator"
ObjTotest(1,0) = "MSWC.BrowserType"
ObjTotest(2,0) = "MSWC.NextLink"
ObjTotest(3,0) = "MSWC.Tools"
ObjTotest(4,0) = "MSWC.Status"
ObjTotest(5,0) = "MSWC.Counters"
ObjTotest(6,0) = "IISSample.ContentRotator"
ObjTotest(7,0) = "IISSample.PageCounter"
ObjTotest(8,0) = "MSWC.PermissionChecker"
ObjTotest(9,0) = "Scripting.FileSystemObject"
ObjTotest(9,1) = "(FSO �ı��ļ���д)"
ObjTotest(10,0) = "adodb.connection"
ObjTotest(10,1) = "(ADO ���ݶ���)"
ObjTotest(11,0) = "SoftArtisans.FileUp"
ObjTotest(11,1) = "(SA-FileUp �ļ��ϴ�)"
ObjTotest(12,0) = "SoftArtisans.FileManager"
ObjTotest(12,1) = "(SoftArtisans �ļ�����)"
ObjTotest(13,0) = "LyfUpload.UploadFile"
ObjTotest(13,1) = "(�ļ��ϴ����)"
ObjTotest(14,0) = "Persits.Upload.1"
ObjTotest(14,1) = "(ASPUpload �ļ��ϴ�)"
ObjTotest(15,0) = "w3.upload"
ObjTotest(15,1) = "(Dimac �ļ��ϴ�)"
ObjTotest(16,0) = "JMail.SmtpMail"
ObjTotest(16,1) = "(Dimac JMail �ʼ��շ�)"
ObjTotest(17,0) = "CDONTS.NewMail"
ObjTotest(17,1) = "(���� SMTP ����)"
ObjTotest(18,0) = "Persits.Mailgraph2er"
ObjTotest(18,1) = "(ASPemail ����)"
ObjTotest(19,0) = "SMTPsvg.Mailer"
ObjTotest(19,1) = "(ASPmail ����)"
ObjTotest(20,0) = "DkQmail.Qmail"
ObjTotest(20,1) = "(dkQmail ����)"
ObjTotest(21,0) = "Geocel.Mailer"
ObjTotest(21,1) = "(Geocel ����)"
ObjTotest(22,0) = "IISmail.Iismail.1"
ObjTotest(22,1) = "(IISmail ����)"
ObjTotest(23,0) = "SmtpMail.SmtpMail.1"
ObjTotest(23,1) = "(SmtpMail ����)"
ObjTotest(24,0) = "SoftArtisans.ImageGen"
ObjTotest(24,1) = "(SA ��ͼ���д���)"
ObjTotest(25,0) = "W3Image.Image"
ObjTotest(25,1) = "(Dimac ��ͼ���д���)"
public IsObj,VerObj

'������֧��������汾

dim i
for i=0 to 25
	on error resume next
	IsObj=false
	VerObj=""
	dim TestObj
	set TestObj=server.CreateObject(ObjTotest(i,0))
	If -2147221005 <> Err then
		IsObj = True
		VerObj = TestObj.version
		if VerObj="" or isnull(VerObj) then VerObj=TestObj.about
	end if
	ObjTotest(i,2)=IsObj
	ObjTotest(i,3)=VerObj
next

'�������Ƿ�֧�ּ�����汾
sub ObjTest(strObj)
	on error resume next
	IsObj=false
	VerObj=""
	dim TestObj
	set TestObj=server.CreateObject (strObj)
	If -2147221005 <> Err then
		IsObj = True
		VerObj = TestObj.version
		if VerObj="" or isnull(VerObj) then VerObj=TestObj.about
	end if	
End sub
%>
<%
if session("admin")="" then
	response.Write("<script>alert('�Ƿ�������');window.location.href='login.asp';</script>")
end if
%>
<style>
BODY {
	font-family: "����";
	font-size: 9pt;
	font-style: normal;
	line-height: 160%;
	color: #000000;
	background-color: #FFFFFF;
}
TABLE {
	font-family: "����";
	font-size: 9pt;
	font-style: normal;
}
A:link {
	FONT-SIZE: 12px; COLOR: #000000; TEXT-DECORATION: none
}
A:visited {
	FONT-SIZE: 12px; COLOR: #000000; TEXT-DECORATION: none
}
A:active {
	FONT-SIZE: 12px; COLOR: #215DC6; TEXT-DECORATION: none
}
A:hover {
	FONT-SIZE: 12px; COLOR: #215DC6; TEXT-DECORATION: none;position: relative; right: 0px; top: 1px
}
.style1 {color: #f2ab5b}
</style>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<table width="100%" border="12" cellpadding="0" cellspacing="0" bordercolor="#6B8FC8" style="border-collapse: collapse">
  <tr align=center bgcolor="6B8FC8" class=backs>
    <td bgcolor="#FFFFFF"><table width="100%" border="0" align="center">
      <tr>
        <td height="555" align="center"><p>���������йز���</p>
            <table width=600 border=0 cellpadding=0 cellspacing=1 bgcolor="#6B8FC8">
              <tr bgcolor="#FFFFFF" height=18>
                <td align=left>&nbsp;��������</td>
                <td>&nbsp;<%=Request.ServerVariables("SERVER_NAME")%></td>
              </tr>
              <tr bgcolor="#FFFFFF" height=18>
                <td align=left>&nbsp;������IP</td>
                <td>&nbsp;<%=Request.ServerVariables("LOCAL_ADDR")%></td>
              </tr>
              <tr bgcolor="#FFFFFF" height=18>
                <td align=left>&nbsp;�������˿�</td>
                <td>&nbsp;<%=Request.ServerVariables("SERVER_PORT")%></td>
              </tr>
              <tr bgcolor="#FFFFFF" height=18>
                <td align=left>&nbsp;������ʱ��</td>
                <td>&nbsp;<%=now%></td>
              </tr>
              <tr bgcolor="#FFFFFF" height=18>
                <td align=left>&nbsp;IIS�汾</td>
                <td>&nbsp;<%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
              </tr>
              <tr bgcolor="#FFFFFF" height=18>
                <td align=left>&nbsp;�ű���ʱʱ��</td>
                <td>&nbsp;<%=Server.ScriptTimeout%> ��</td>
              </tr>
              <tr bgcolor="#FFFFFF" height=18>
                <td align=left>&nbsp;���ļ�·��</td>
                <td>&nbsp;<%=server.mappath(Request.ServerVariables("SCRIPT_NAME"))%></td>
              </tr>
              <tr bgcolor="#FFFFFF" height=18>
                <td align=left>&nbsp;������CPU����</td>
                <td>&nbsp;<%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%> ��</td>
              </tr>
              <tr bgcolor="#FFFFFF" height=18>
                <td align=left>&nbsp;��������������</td>
                <td>&nbsp;<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
              </tr>
              <tr bgcolor="#FFFFFF" height=18>
                <td align=left>&nbsp;����������ϵͳ</td>
                <td>&nbsp;<%=Request.ServerVariables("OS")%></td>
              </tr>
            </table>
            <br>
            <font class=fonts>���֧�����</font> <br>
            <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#6B8FC8" width="600">
              <tr align=center bgcolor="6B8FC8" class=backs height=18>
                <td width=320><font color="#FFFFFF">�� �� �� ��</font></td>
                <td width=130><font color="#FFFFFF">֧�ּ��汾</font></td>
              </tr>
              <%For i=0 to 10%>
              <tr height="18" class=backq>
                <td align=left>&nbsp;<%=ObjTotest(i,0) & "<font color=#888888>&nbsp;" & ObjTotest(i,1)%></td>
                <td align=left>&nbsp;
                    <%
		If Not ObjTotest(i,2) Then 
			Response.Write "<font color=red><b>��</b></font>"
		Else
			Response.Write "<font class=fonts><b>��</b></font> <a title='" & ObjTotest(i,3) & "'>" & left(ObjTotest(i,3),11) & "</a>"
		End If%>
                </td>
              </tr>
              <%next%>
          </table></td>
      </tr>
    </table></td>
  </tr>
</table>
</BODY>
</HTML>