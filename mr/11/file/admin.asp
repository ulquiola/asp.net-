<!--#include file="Conn.asp"-->
<!--#include file="Function.asp"-->
<!--#include file=inc.asp-->
<%
Dim Rs,Sql
dim GroupID,GroupName,UpFileType,UpFilesize
Dim UserName,UserPwd,UserSex,Company,Department,Tel,Mobile,UserType,IP
If Session("UserName")="" Then
Response.Write("<script language=JavaScript>alert('����û�е�½��');document.location.href='index.asp';</script>")
Response.End
End if
If Not Session("UserAdmin") Then
Response.Write("<script language=JavaScript>alert('��û��Ȩ�޲�����');document.location.href='javascript:window.history.go(-1);';</script>")
Response.End
End If
IF Request.QueryString("Action")="Savead" Then
If ChkPost=false Then
Response.Write("<script language=JavaScript>alert('�벻Ҫ���ⲿ�ύ��Ϣ��');document.location.href='index.asp';</script>")
Response.End
End If
Sql="Select * From tb_Setup"
Set Rs=server.CreateObject("adodb.recordset")
If Not IsObject(Conn) Then ConnectionDatabase()
Rs.open Sql,Conn,1,3
if rs.eof then
rs.addnew
rs("Info")=CheckStr(Request.Form("Info"))
rs.update
else
rs("Info")=CheckStr(Request.Form("Info"))
rs.update
end if
rs.close
set rs=nothing
response.Redirect("Admin.asp?action=ad")
end if

IF Request.QueryString("Action")="delfile" Then
If ChkPost=false Then
Response.Write("<script language=JavaScript>alert('�벻Ҫ���ⲿ�ύ��Ϣ��');document.location.href='index.asp';</script>")
Response.End
End If

dim id
Dim myobjFSO
id=Request.QueryString("id")
if id="" or not isnumeric(id) then
Response.Write("<script language=JavaScript>alert('���������ʧ��');document.location.href='javascript:window.history.go(-1);';</script>")
Response.End
end if
id=clng(id)
Sql="Select * From tb_Files Where id="&id&""
Set Rs=server.CreateObject("adodb.recordset")
If Not IsObject(Conn) Then ConnectionDatabase()
Rs.open Sql,Conn,1,3
If Rs.Eof Then
Rs.close
Set Rs=Nothing
CloseDatabase()
Response.Write("<script language=JavaScript>alert('�ļ������ڣ�');document.location.href='javascript:window.history.go(-1);';</script>")
response.End
Else
Set myobjFSO = Server.CreateObject("Scripting.FileSystemObject")
if myobjfso.FileExists(Server.MapPath("UpLoadFile/"&rs("FileUrl"))) then
myobjFSO.DeleteFile(Server.MapPath("UpLoadFile/"&rs("FileUrl")))
end if
set myobjFSO=nothing
Conn.Execute("Update tb_Users set UpFiles=UpFiles-1 Where UserName='"&rs("UserName")&"'")
Rs.delete
end if
rs.close
set rs=nothing
CloseDatabase()
Response.Write("<script language=JavaScript>alert('ɾ���ɹ���');document.location.href='Admin.asp?action=file';</script>")
end if

IF Request.QueryString("Action")="addgroup" Then
If ChkPost=false Then
Response.Write("<script language=JavaScript>alert('�벻Ҫ���ⲿ�ύ��Ϣ��');document.location.href='index.asp';</script>")
Response.End
End If

GroupName=Trim(Request.Form("GroupName"))
UpFileType=Trim(Request.Form("UpFileType"))
UpFilesize=Trim(Request.Form("UpFilesize"))
if GroupName="" then
Response.Write("<script language=JavaScript>alert('������������ƣ�');document.location.href='javascript:window.history.go(-1);';</script>")
Response.End
End If
GroupName=Checkstr(GroupName)
UpFileType=Checkstr(UpFileType)
UpFilesize=Checkstr(UpFilesize)
if UpFilesize<>"" then
if isInteger(UpFilesize)=false then
Response.Write("<script language=JavaScript>alert('�����ϴ����ļ���С������������');document.location.href='javascript:window.history.go(-1);';</script>")
Response.End
end if
else
UpFilesize=0
end if

Sql="Select * From tb_Groups Where GroupName='"&GroupName&"'"
Set Rs=server.CreateObject("adodb.recordset")
If Not IsObject(Conn) Then ConnectionDatabase()
Rs.open Sql,Conn,1,3
If Not Rs.Eof Then
Rs.close
Set Rs=Nothing
CloseDatabase()
Response.Write("<script language=JavaScript>alert('�������Ѿ����ڣ�');document.location.href='javascript:window.history.go(-1);';</script>")
Else
Rs.AddNew
Rs("GroupName")=GroupName
Rs("UpFileType")=UpFileType
Rs("UpFilesize")=UpFilesize
Rs.UpDate
Rs.close
Set Rs=nothing
CloseDatabase()
Response.Write("<script language=JavaScript>alert('�����ɹ���');document.location.href='Admin.asp?action=group';</script>")
End if
Response.End
End If

IF Request.QueryString("Action")="editgroup" Then
If ChkPost=false Then
Response.Write("<script language=JavaScript>alert('�벻Ҫ���ⲿ�ύ��Ϣ��');document.location.href='index.asp';</script>")
Response.End
End If

GroupID=Trim(Request.Form("GroupID"))
GroupName=Trim(Request.Form("GroupName"))
UpFileType=Trim(Request.Form("UpFileType"))
UpFilesize=Trim(Request.Form("UpFilesize"))
IF GroupID="" OR NOT ISnumeric(GroupID) then
Response.Write("<script language=JavaScript>alert('���������ʧ��');document.location.href='javascript:window.history.go(-1);';</script>")
Response.End
else
GroupID=clng(GroupID)
end if

if GroupName="" then
Response.Write("<script language=JavaScript>alert('������������ƣ�');document.location.href='javascript:window.history.go(-1);';</script>")
Response.End
End If
GroupName=Checkstr(GroupName)
UpFileType=Checkstr(UpFileType)
UpFilesize=Checkstr(UpFilesize)
if UpFilesize<>"" then
if isInteger(UpFilesize)=false then
Response.Write("<script language=JavaScript>alert('�����ϴ����ļ���С������������');document.location.href='javascript:window.history.go(-1);';</script>")
Response.End
end if
else
UpFilesize=0
end if

Sql="Select * From tb_Groups Where GroupID="&GroupID&""
Set Rs=server.CreateObject("adodb.recordset")
If Not IsObject(Conn) Then ConnectionDatabase()
Rs.open Sql,Conn,1,3
If Rs.Eof Then
Rs.close
Set Rs=Nothing
CloseDatabase()
Response.Write("<script language=JavaScript>alert('���鲻���ڣ�');document.location.href='javascript:window.history.go(-1);';</script>")
Else
Rs("GroupName")=GroupName
Rs("UpFileType")=UpFileType
Rs("UpFilesize")=UpFilesize
Rs.UpDate
Rs.close
Set Rs=nothing
CloseDatabase()
Response.Write("<script language=JavaScript>alert('�޸���ɹ���');document.location.href='Admin.asp?action=group';</script>")
End if
Response.End
End If

IF Request.QueryString("Action")="delgroup" Then

GroupID=Trim(Request.QueryString("groupid"))

IF GroupID="" OR NOT ISnumeric(GroupID) then
Response.Write("<script language=JavaScript>alert('���������ʧ��');document.location.href='javascript:window.history.go(-1);';</script>")
Response.End
else
GroupID=clng(GroupID)
end if

Sql="Select * From tb_Groups Where GroupID="&GroupID&""
Set Rs=server.CreateObject("adodb.recordset")
If Not IsObject(Conn) Then ConnectionDatabase()
Rs.open Sql,Conn,1,3
If Rs.Eof Then
Rs.close
Set Rs=Nothing
CloseDatabase()
Response.Write("<script language=JavaScript>alert('���鲻���ڣ�');document.location.href='javascript:window.history.go(-1);';</script>")
Else
if Conn.Execute("Select count(*) from tb_Users where GroupID="&GroupID&"")(0)>0 then
Rs.close
Set Rs=nothing
CloseDatabase()
Response.Write("<script language=JavaScript>alert('�����Ѿ����û��������Ƴ������û���');document.location.href='javascript:window.history.go(-1);';</script>")
response.End
else
Rs.delete
end if
Rs.close
Set Rs=nothing
CloseDatabase()
Response.Write("<script language=JavaScript>alert('ɾ����ɹ���');document.location.href='Admin.asp?action=group';</script>")
End if
Response.End
End If

If Request.QueryString("Action")="adduser" Then
If ChkPost=false Then
Response.Write("<script language=JavaScript>alert('�벻Ҫ���ⲿ�ύ��Ϣ��');document.location.href='index.asp';</script>")
Response.End
End If

UserName=Trim(Request.Form("UserName"))
UserPwd=Trim(Request.Form("UserPwd"))
UserSex=Checkstr(Trim(Request.Form("UserSex")))
Company=Checkstr(Trim(Request.Form("Company")))
Department=Checkstr(Trim(Request.Form("Department")))
Tel=Checkstr(Trim(Request.Form("Tel")))
Mobile=Checkstr(Trim(Request.Form("Mobile")))
UserType=Checkstr(Trim(Request.Form("UserType")))
GroupID=Checkstr(Trim(Request.Form("GroupID")))
If UserName="" or UserPwd="" Then
Response.Write("<script language=JavaScript>alert('�������û������û����룡');document.location.href='javascript:window.history.go(-1);';</script>")
Response.End
End If
UserName=Checkstr(UserName)

if Tel<>"" Then
if IsValidTel(Tel)=False then
Response.Write("<script language=JavaScript>alert('�绰�����ʽ����');document.location.href='javascript:window.history.go(-1);';</script>")
Response.End
End If
end if

if Mobile<>"" then
If isInteger(Mobile)=False Then
Response.Write("<script language=JavaScript>alert('�ֻ�������������֣�');document.location.href='javascript:window.history.go(-1);';</script>")
Response.End
End If
end if

IF GroupID="" or GroupID="0" OR NOT ISnumeric(GroupID) then
Response.Write("<script language=JavaScript>alert('����ȷѡ���û��飡');document.location.href='javascript:window.history.go(-1);';</script>")
Response.End
else
GroupID=clng(GroupID)
end if

Sql="Select * From tb_Users Where UserName='"&UserName&"'"
Set Rs=server.CreateObject("adodb.recordset")
If Not IsObject(Conn) Then ConnectionDatabase()
Rs.open Sql,Conn,1,3
If Not Rs.Eof Then
Rs.close
Set Rs=Nothing
CloseDatabase()
Response.Write("<script language=JavaScript>alert('���û����Ѿ����ڣ�');document.location.href='javascript:window.history.go(-1);';</script>")
Else
Rs.AddNew
If Request.ServerVariables("HTTP_X_FORWARDED_FOR")<>"" then 
IP=Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
Else
IP=Request.ServerVariables("REMOTE_ADDR")
End If
Rs("UserName")=UserName
Rs("UserPwd")=UserPwd
Rs("UserSex")=Cbool(UserSex)
Rs("Company")=Company
Rs("Department")=Department
Rs("Tel")=Tel
Rs("Mobile")=Mobile
Rs("UserRegIp")=IP
Rs("UserLastIp")=IP
Rs("GroupID")=GroupID
Rs("UserType")=CBool(UserType)
Rs.UpDate
Rs.close
Set Rs=nothing
CloseDatabase()
Response.Write("<script language=JavaScript>alert('�û���ӳɹ���');document.location.href='Admin.asp';</script>")
End if
Response.End
End If

If Request.QueryString("Action")="saveuser" Then
If ChkPost=false Then
Response.Write("<script language=JavaScript>alert('�벻Ҫ���ⲿ�ύ��Ϣ��');document.location.href='Index.asp';</script>")
Response.End
End If

UserName=Trim(Request.Form("UserName"))
UserPwd=Trim(Request.Form("UserPwd"))
UserSex=Checkstr(Trim(Request.Form("UserSex")))
Company=Checkstr(Trim(Request.Form("Company")))
Department=Checkstr(Trim(Request.Form("Department")))
Tel=Checkstr(Trim(Request.Form("Tel")))
Mobile=Checkstr(Trim(Request.Form("Mobile")))
UserType=Checkstr(Trim(Request.Form("UserType")))
GroupID=Checkstr(Trim(Request.Form("GroupID")))

If UserName="" Then
Response.Write("<script language=JavaScript>alert('������ʧ��');document.location.href='javascript:window.history.go(-1);';</script>")
Response.End
End If
UserName=Checkstr(UserName)

if Tel<>"" Then
if IsValidTel(Tel)=False then
Response.Write("<script language=JavaScript>alert('�绰�����ʽ����');document.location.href='javascript:window.history.go(-1);';</script>")
Response.End
End If
end if

if Mobile<>"" then
If isInteger(Mobile)=False Then
Response.Write("<script language=JavaScript>alert('�ֻ�������������֣�');document.location.href='javascript:window.history.go(-1);';</script>")
Response.End
End If
end if

IF GroupID="" or GroupID="0" OR NOT ISnumeric(GroupID) then
Response.Write("<script language=JavaScript>alert('����ȷѡ���û��飡');document.location.href='javascript:window.history.go(-1);';</script>")
Response.End
else
GroupID=clng(GroupID)
end if

Sql="Select * From tb_Users Where UserName='"&UserName&"'"
Set Rs=server.CreateObject("adodb.recordset")
If Not IsObject(Conn) Then ConnectionDatabase()
Rs.open Sql,Conn,1,3
If Rs.Eof Then
Rs.close
Set Rs=Nothing
CloseDatabase()
Response.Write("<script language=JavaScript>alert('���û������ڣ�');document.location.href='javascript:window.history.go(-1);';</script>")
Else

Rs("UserPwd")=UserPwd
Rs("UserSex")=Cbool(UserSex)
Rs("Company")=Company
Rs("Department")=Department
Rs("Tel")=Tel
Rs("Mobile")=Mobile
Rs("GroupID")=GroupID
Rs("UserType")=CBool(UserType)
Rs.UpDate
Rs.close
Set Rs=nothing
CloseDatabase()
Response.Write("<script language=JavaScript>alert('�����޸ĳɹ���');document.location.href='Admin.asp';</script>")
End if
Response.End
End If


If Request.QueryString("Action")="deluser" Then
If ChkPost=false Then
Response.Write("<script language=JavaScript>alert('�벻Ҫ���ⲿ�ύ��Ϣ��');document.location.href='index.asp';</script>")
Response.End
End If
Dim objFSO,rs2
UserName=Checkstr(trim(request.QueryString("UserName")))
if UserName="" Then
Response.Write("<script language=JavaScript>alert('������ʧ��');document.location.href='javascript:window.history.go(-1);';</script>")
else
Sql="Select * From tb_Users where UserName='"&UserName&"'"
set Rs=server.CreateObject("adodb.recordset")
If Not IsObject(Conn) Then ConnectionDatabase()
Rs.open sql,conn,1,3
if rs.eof then
Rs.close
Set Rs=nothing
CloseDatabase()
Response.Write("<script language=JavaScript>alert('���û������ڣ�');document.location.href='javascript:window.history.go(-1);';</script>")
Response.End
else
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
set Rs2=server.CreateObject("adodb.recordset")
Sql="Select * From tb_Files where UserName='"&UserName&"'"
Rs2.open sql,conn,1,3
do until rs2.eof
if objfso.FileExists(Server.MapPath("UpLoadFile/"&rs2("FileUrl"))) then
objFSO.DeleteFile(Server.MapPath("UpLoadFile/"&rs2("FileUrl")))
end if
Rs2.delete
rs2.movenext
loop
rs2.close
set rs2=nothing
Set objFSO =nothing
Rs.delete
rs.close
set rs=nothing
CloseDatabase()
Response.Write("<script language=JavaScript>alert('ɾ���ɹ���');document.location.href='Admin.asp';</script>")
End if
End if
Response.End
End if
%>
<%Call Top("ϵͳ����","Admin.asp")%>
<TABLE width=785 align="center">
  <TBODY>
  <TR>
    <TD style="FONT-SIZE: 12px; WIDTH: 200px; HEIGHT: 21px; TEXT-ALIGN: center" 
    align=left rowSpan=2 valign="top">
      <TABLE class="table">
        <TBODY>
        <TR>
          <TD width="200" 
          colSpan=3 bgcolor="#3399ff" style="FONT-WEIGHT: bold; WIDTH: 200px; COLOR: white; HEIGHT: 25px; TEXT-ALIGN: center; TEXT-DECORATION: none">�û���½</TD>
        </TR>
        <TR>
          <TD style="WIDTH: 200px; HEIGHT: 145px; TEXT-ALIGN: center" 
          colSpan=3 rowSpan=2>
            <TABLE style="WIDTH: 200px; HEIGHT: 126px; TEXT-ALIGN: left">
              <TBODY>
              <TR>
                <TD style="WIDTH: 188px" colSpan=3 rowSpan=3><%Call Login()%></TD></TR>
              <TR></TR>
              <TR></TR></TBODY></TABLE></TD></TR>
        <TR></TR></TBODY></TABLE></TD>
    <TD style="WIDTH:100%; HEIGHT: 21px" align=center valign="top" colSpan=2 rowSpan=2>
      <TABLE style="WIDTH: 100%; HEIGHT: 182px">
        <TBODY>
        <TR>
          <TD height="17" colSpan=3 class="titletd" style="FONT-WEIGHT: bold;COLOR: blue;TEXT-ALIGN: center;">
		  <a href="Admin.asp">�û��б�</a>��<a href="Admin.asp?action=add">����û�</a>��<a href="Admin.asp?action=group">�û������</a>��<a href="Admin.asp?action=file">�ļ�����</a>��<a href="Admin.asp?action=ad">��ҳ����</a></TD>
        </TR>
        <TR>
          <TD class="table" colSpan=3 rowSpan=3 valign="top" align="center">
		  <%
if Request.QueryString("Action")="add" then
%>
<table width="100%" border="0" cellspacing="3" cellpadding="0">
<form id="form1" name="form1" method="post" action="Admin.asp?action=adduser">
              <tr>
                <td width="21%" align="right">�û�����</td>
                <td width="79%"><input name="UserName" type="text" id="UserName" maxlength="20" />&nbsp;��Ӻ󲻿��޸ġ�</td>
              </tr>
              <tr>
                <td  align="right">�ܡ��룺</td>
                <td><input name="UserPwd" type="text" id="UserPwd" maxlength="20" /></td>
              </tr>
              <tr>
                <td  align="right">�ԡ���</td>
                <td><input type="radio" name="UserSex" value="1" checked />
                  ��
                    <input type="radio" name="UserSex" value="0" />
                    Ů</td>
              </tr>
              <tr>
                <td align="right">����λ��</td>
                <td><input name="Company" type="text" id="Company" maxlength="48" /></td>
              </tr>
              <tr>
                <td align="right">�����ţ�</td>
                <td><input name="Department" type="text" id="Department" maxlength="48" /></td>
              </tr>
              <tr>
                <td align="right">�֡�����</td>
                <td><input name="Mobile" type="text" id="Mobile" maxlength="20" /></td>
              </tr>
              <tr>
                <td align="right">�硡����</td>
                <td><input name="Tel" type="text" id="Tel" maxlength="20" /></td>
              </tr>
              <tr>
                <td align="right">����Ա��</td>
                <td><input name="UserType" type="radio" value="1" />
                  �ǹ���Ա
                    <input name="UserType" type="radio" value="0" checked="checked" />
                    �ǹ���Ա</td>
              </tr>
              <tr>
                <td align="right">�����飺</td>
                <td>
				<%If Not IsObject(Conn) Then ConnectionDatabase()
set rs=Conn.execute("Select * from tb_Groups")
if rs.eof then
response.Write("��������飡")
else
response.Write("<select name=""GroupID"" id=""GroupID"">")
response.Write("<option value=""0"">��ѡ���û���</option>")
do until rs.eof
response.Write("<option value="""&Rs("GroupID")&""">"&Rs("GroupName")&"</option>")
rs.movenext
loop
response.Write("</select>")
end if
set rs=nothing
%>
                </td>
              </tr>
              <tr>
                <td colspan="2" align="center"><input type="submit" name="Submit" value="���" /></td>
                </tr>
				</form>
            </table>
<%
elseif Request.QueryString("Action")="edituser" then
If Not IsObject(Conn) Then ConnectionDatabase()
		  Set Rs=Conn.execute("Select * from tb_Users Where UserName='"&Checkstr(Request.QueryString("UserName"))&"'")
		  if rs.eof then
		   set Rs=nothing
		   CloseDatabase()
		Response.Write("<script language=JavaScript>alert('�û������ڣ������µ�½��');document.location.href='javascript:window.history.go(-1);';</script>")
		Response.End
		  End if
		  GroupID=rs("GroupID")
%>
<table width="100%" border="0" cellspacing="3" cellpadding="0">
<form id="form1" name="form1" method="post" action="Admin.asp?action=saveuser">
              <tr>
                <td width="21%" align="right">�û�����</td>
                <td width="79%"><%=rs("UserName")%><input name="UserName" type="hidden" id="UserName" value="<%=rs("UserName")%>" /></td>
              </tr>
              <tr>
                <td  align="right">�ܡ��룺</td>
                <td><input name="UserPwd" type="text" id="UserPwd" maxlength="20" />&nbsp;���޸��벻Ҫ��д��</td>
              </tr>
              <tr>
                <td  align="right">�ԡ���</td>
                <td><input type="radio" name="UserSex" value="1" <%if rs("UserSex") then response.Write("checked")%> />
                  ��
                    <input type="radio" name="UserSex" value="0" <%if not rs("UserSex") then response.Write("checked")%> />
                    Ů</td>
              </tr>
              <tr>
                <td align="right">����λ��</td>
                <td><input name="Company" type="text" id="Company" value="<%=rs("Company")%>" maxlength="48" /></td>
              </tr>
              <tr>
                <td align="right">�����ţ�</td>
                <td><input name="Department" type="text" id="Department" value="<%=rs("Department")%>" maxlength="48" /></td>
              </tr>
              <tr>
                <td align="right">�֡�����</td>
                <td><input name="Mobile" type="text" id="Mobile" value="<%=rs("Mobile")%>" maxlength="20" /></td>
              </tr>
              <tr>
                <td align="right">�硡����</td>
                <td><input name="Tel" type="text" id="Tel" value="<%=rs("Tel")%>" maxlength="20" /></td>
              </tr>
              <tr>
                <td align="right">����Ա��</td>
                <td><input name="UserType" type="radio" value="1" <%if rs("UserType") then response.Write("checked")%> />
                  �ǹ���Ա
                    <input name="UserType" type="radio" value="0" <%if not rs("UserType") then response.Write("checked")%> />
                    �ǹ���Ա</td>
              </tr>
              <tr>
                <td align="right">�����飺</td>
                <td>
				<%If Not IsObject(Conn) Then ConnectionDatabase()
set rs=Conn.execute("Select * from tb_Groups")
if rs.eof then
response.Write("��������飡")
else
response.Write("<select name=""GroupID"" id=""GroupID"">")
response.Write("<option value=""0"">��ѡ���û���</option>")
do until rs.eof
response.Write("<option value="""&Rs("GroupID")&"""")
if GroupID=Rs("GroupID") then response.Write(" selected ")
response.Write(">"&Rs("GroupName")&"</option>")
rs.movenext
loop
response.Write("</select>")
end if
set rs=nothing
%>
                  
               
                </td>
              </tr>
              <tr>
                <td colspan="2" align="center"><input type="submit" name="Submit" value="�޸�" /></td>
                </tr>
				</form>
            </table>
<%
elseif Request.QueryString("Action")="file" then
Call Filelist()
elseif Request.QueryString("Action")="ad" then
dim adStr
set rs=conn.execute("select Info from tb_Setup")
if not rs.eof then
adStr=rs(0)
end if
set rs=nothing
response.Write("<form id='form1' name='form1' method='post' action='Admin.asp?action=Savead'>")
response.Write("<textarea name='Info' cols='60' rows='10' id='Info'>"&adStr&"</textarea>")
response.Write("<br><input type='submit' name='Submit' value='�ύ' />")
response.Write("</form>")
elseif Request.QueryString("Action")="group" then
response.Write("<table width=""100%"" border=""0"" cellspacing=""1"" cellpadding=""3"">"&_
               "<tr align=""center"" bgcolor=""#3399ff"">"&_
               "<td width=""20%"" style=""FONT-WEIGHT: bold; COLOR: white; HEIGHT: 25px; TEXT-ALIGN: center; TEXT-DECORATION: none"">����</td>"&_
               "<td width=""30%"" style=""FONT-WEIGHT: bold; COLOR: white; HEIGHT: 25px; TEXT-ALIGN: center; TEXT-DECORATION: none"">�����ϴ��ļ�����</td>"&_
               "<td width=""30%"" style=""FONT-WEIGHT: bold; COLOR: white; HEIGHT: 25px; TEXT-ALIGN: center; TEXT-DECORATION: none"">�����ϴ��ļ���С</td>"&_
               "<td width=""20%"" style=""FONT-WEIGHT: bold; COLOR: white; HEIGHT: 25px; TEXT-ALIGN: center; TEXT-DECORATION: none"">����</td>"&_
               "</tr>")
If Not IsObject(Conn) Then ConnectionDatabase()
set rs=Conn.execute("Select * from tb_Groups")
if rs.eof then
response.write"<tr><td colspan=4>��û���κ��飬����ӡ���</td></tr>"
else
do until rs.eof
response.Write("<form id=""form"&rs("GroupID")&""" name=""form"&rs("GroupID")&""" method=""post"" action=""Admin.asp?action=editgroup""><input name=""GroupID"" type=""hidden"" id=""GroupID"" value="""&rs("GroupID")&""" /><tr class=bgcolor_1><td><input name=""GroupName"" type=""text"" id=""GroupName"" value="""&rs("GroupName")&""" size=""15"" maxlength=""20"" /></td><td><input name=""UpFileType"" type=""text"" value="""&rs("UpFileType")&""" id=""UpFileType"" size=""20"" maxlength=""150"" /></td><TD><input name=""UpFilesize"" type=""text"" id=""UpFilesize"" size=""10"" value="""&rs("UpFilesize")&""" maxlength=""15"" />&nbsp;KB</TD><td><input type=""submit"" name=""Submit"" value=""�޸�"" />&nbsp;<input name=""DEL"" type=""button"" id=""DEL"" value=""ɾ��""  onClick=""location.href='Admin.asp?action=delgroup&groupid="&rs("GroupID")&"'"" /></TD></tr></form>")
rs.movenext
loop
end if
set rs=nothing
response.Write("<form id=""formadd"" name=""formadd"" method=""post"" action=""Admin.asp?action=addgroup""><tr class=bgcolor_1><td><input name=""GroupName"" type=""text"" id=""GroupName"" size=""15"" maxlength=""20"" /></td><td><input name=""UpFileType"" type=""text"" id=""UpFileType"" size=""20"" maxlength=""150"" /></td><TD><input name=""UpFilesize"" type=""text"" id=""UpFilesize"" size=""10"" value=""0"" maxlength=""15"" />&nbsp;KB</TD><td><input type=""submit"" name=""Submit"" value=""���"" /></TD></tr></form>")	
response.write"<tr><td colspan=4>˵���������ϴ������ļ�������ʹ�á�|���ָ����磺RAR|ZIP|DOC�������ϴ����ļ���С���������������������ϴ������ա�</td></tr>"   
response.Write("</table>")     
else
response.Write("<table width=""100%"" border=""0"" cellspacing=""1"" cellpadding=""3"">"&_
               "<tr align=""center"" bgcolor=""#3399ff"">"&_
               "<td width=""12%"" style=""FONT-WEIGHT: bold; COLOR: white; HEIGHT: 25px; TEXT-ALIGN: center; TEXT-DECORATION: none"">�û���</td>"&_
               "<td width=""6%"" style=""FONT-WEIGHT: bold; COLOR: white; HEIGHT: 25px; TEXT-ALIGN: center; TEXT-DECORATION: none"">�Ա�</td>"&_
               "<td width=""15%"" style=""FONT-WEIGHT: bold; COLOR: white; HEIGHT: 25px; TEXT-ALIGN: center; TEXT-DECORATION: none"">��λ</td>"&_
               "<td width=""12%"" style=""FONT-WEIGHT: bold; COLOR: white; HEIGHT: 25px; TEXT-ALIGN: center; TEXT-DECORATION: none"">����</td>"&_
               "<td width=""16%"" style=""FONT-WEIGHT: bold; COLOR: white; HEIGHT: 25px; TEXT-ALIGN: center; TEXT-DECORATION: none"">�绰</td>"&_
               "<td width=""10%"" style=""FONT-WEIGHT: bold; COLOR: white; HEIGHT: 25px; TEXT-ALIGN: center; TEXT-DECORATION: none""><a href=Admin.asp?orderby=group title=��������><font color=white>������</font></a></td>"&_
               "<td width=""10%"" style=""FONT-WEIGHT: bold; COLOR: white; HEIGHT: 25px; TEXT-ALIGN: center; TEXT-DECORATION: none"">����</td>"&_
               "<td width=""6%"" style=""FONT-WEIGHT: bold; COLOR: white; HEIGHT: 25px; TEXT-ALIGN: center; TEXT-DECORATION: none"">�ϴ�</td>"&_
               "<td width=""13%"" style=""FONT-WEIGHT: bold; COLOR: white; HEIGHT: 25px; TEXT-ALIGN: center; TEXT-DECORATION: none"">����</td>"&_
               "</tr>")
		  dim totalPut
          dim Num,orderby,sqlStr,Gid
          dim strFileName,currentPage,MaxPerPage
		  orderby=Trim(Request("orderby"))
		  Gid=Trim(Request("GroupID"))
		   sqlStr=""
		   strFileName="Admin.asp?"
		  if orderby="group" then
		     sqlStr="order by GroupID"
			 strFileName="Admin.asp?orderby=group&"
		  end if
		  if Gid<>"" and isnumeric(Gid) then
		     sqlStr="where GroupID="&Gid&""
			 strFileName="Admin.asp?GroupID="&Gid&"&"
		  end if
		  
          MaxPerPage=10
          if request("page")<>"" then
             currentPage=clng(request("page"))
          else
	         currentPage=1
          end if
		  If Not IsObject(Conn) Then ConnectionDatabase()
		  Set rs= Server.CreateObject("ADODB.Recordset")
		  sql="Select * from tb_Users "&sqlStr&" "
		  rs.open sql,conn,1,1
		  if  rs.eof then
			 response.write"<tr><td colspan=9>��û���κ��û�����</td></tr>"
		  else
		       totalPut=rs.recordcount
	        if currentpage<1 then
   		       currentpage=1
   	        end if
   	           if (currentpage-1)*MaxPerPage>totalput then
   		       if (totalPut mod MaxPerPage)=0 then
     		      currentpage= totalPut \ MaxPerPage
	  	       else
	      	      currentpage= totalPut \ MaxPerPage + 1
   		       end if
   	           end if
		if currentPage=1 then
	    Num=0
       	do while not rs.eof
		response.Write("<tr class=bgcolor_1><td>"&rs("UserName")&"</td><td align=center>")
		if rs("UserSex") then response.Write("��") else response.Write("Ů")
		response.Write("</td><td>"&rs("Company")&"</td><td>"&rs("Department")&"</td><td>"&rs("Tel")&"</td><td><a href=Admin.asp?GroupID="&Rs("GroupID")&" title=�鿴�����û�>"&GetGroupInfo(rs("GroupID"))(0)&"</a></td><td>")
		if rs("UserType") then response.Write("����Ա") else response.Write("��ͨ�û�")
		response.Write("</td><td align=center>"&rs("UpFiles")&"</td><td><a href=Admin.asp?action=deluser&UserName="&rs("UserName")&" onClick=""return confirm('��ȷ��Ҫִ��ɾ������?')"">ɾ��</a>��<a href=Admin.asp?action=edituser&UserName="&rs("UserName")&">�޸�</a></td></tr>")
       	Num=Num+1
	   	if Num>=MaxPerPage then exit do
	   	rs.movenext
	    loop
		
 	else
     	if (currentPage-1)*MaxPerPage<totalPut then
       	   	rs.move  (currentPage-1)*MaxPerPage
       		dim bookmark
           	bookmark=rs.bookmark
        Num=0
       	do while not rs.eof
		response.Write("<tr class=bgcolor_1><td>"&rs("UserName")&"</td><td align=center>")
		if rs("UserSex") then response.Write("��") else response.Write("Ů")
		response.Write("</td><td>"&rs("Company")&"</td><td>"&rs("Department")&"</td><td>"&rs("Tel")&"</td><td>"&GetGroupInfo(rs("GroupID"))(0)&"</td><td>")
		if rs("UserType") then response.Write("����Ա") else response.Write("��ͨ�û�")
		response.Write("</td><td align=center>"&rs("UpFiles")&"</td><td><a href=Admin.asp?action=deluser&UserName="&rs("UserName")&" onClick=""return confirm('��ȷ��Ҫִ��ɾ������?')"">ɾ��</a>��<a href=Admin.asp?action=edituser&UserName="&rs("UserName")&">�޸�</a></td></tr>")
       	Num=Num+1
	   	if Num>=MaxPerPage then exit do
	   	rs.movenext
	    loop
          
       	else
	        currentPage=1
        Num=0
       	do while not rs.eof
		response.Write("<tr class=bgcolor_1><td>"&rs("UserName")&"</td><td align=center>")
		if rs("UserSex") then response.Write("��") else response.Write("Ů")
		response.Write("</td><td>"&rs("Company")&"</td><td>"&rs("Department")&"</td><td>"&rs("Tel")&"</td><td>"&GetGroupInfo(rs("GroupID"))(0)&"</td><td>")
		if rs("UserType") then response.Write("����Ա") else response.Write("��ͨ�û�")
		response.Write("</td><td align=center>"&rs("UpFiles")&"</td><td><a href=Admin.asp?action=deluser&UserName="&rs("UserName")&" onClick=""return confirm('��ȷ��Ҫִ��ɾ������?')"">ɾ��</a>��<a href=Admin.asp?action=edituser&UserName="&rs("UserName")&">�޸�</a></td></tr>")
       	Num=Num+1
	   	if Num>=MaxPerPage then exit do
	   	rs.movenext
	    loop
          	
	    end if
	end if
	response.Write("<tr><td colspan=9>"&showpage(strFileName,totalput,MaxPerPage,CurrentPage,true,true,"���û�")&"</td></tr>")
end if
rs.close
set rs=nothing
response.Write("</table>")
end if
%>
</TD>
        </TR>
        <TR></TR>
        <TR></TR>
        <TR>
          <TD class="titletd" colSpan=3>&nbsp;</TD>
        </TR></TBODY></TABLE></TD></TR>
  <TR></TR></TBODY></TABLE>
  <%Sub Filelist()
   response.Write("<table width=""100%"" border=""0"" cellspacing=""1"" cellpadding=""3"">"&_
               "<tr align=""center"" bgcolor=""#3399ff"">"&_
               "<td width=""25%"" style=""FONT-WEIGHT: bold; COLOR: white; HEIGHT: 25px; TEXT-ALIGN: center; TEXT-DECORATION: none"">�ļ���</td>"&_
               "<td width=""10%"" style=""FONT-WEIGHT: bold; COLOR: white; HEIGHT: 25px; TEXT-ALIGN: center; TEXT-DECORATION: none"">�ϴ���</td>"&_
               "<td width=""10%"" style=""FONT-WEIGHT: bold; COLOR: white; HEIGHT: 25px; TEXT-ALIGN: center; TEXT-DECORATION: none"">�ļ���С</td>"&_
               "<td width=""18%"" style=""FONT-WEIGHT: bold; COLOR: white; HEIGHT: 25px; TEXT-ALIGN: center; TEXT-DECORATION: none"">�ϴ�ʱ��</td>"&_
               "<td width=""10%"" style=""FONT-WEIGHT: bold; COLOR: white; HEIGHT: 25px; TEXT-ALIGN: center; TEXT-DECORATION: none"">�ϴ�IP</td>"&_
               "<td width=""6%"" style=""FONT-WEIGHT: bold; COLOR: white; HEIGHT: 25px; TEXT-ALIGN: center; TEXT-DECORATION: none"">����</td>"&_
               "</tr>")
		  dim totalPut
          dim Num
          dim strFileName,currentPage,MaxPerPage
		  strFileName="Admin.asp?action=file&"
          MaxPerPage=10
          if request("page")<>"" then
             currentPage=clng(request("page"))
          else
	         currentPage=1
          end if
		  If Not IsObject(Conn) Then ConnectionDatabase()
		  Set rs= Server.CreateObject("ADODB.Recordset")
		  sql="Select * from tb_Files order by id desc "
		  rs.open sql,conn,1,1
		  if  rs.eof then
			 response.write"<tr><td colspan=6>��û���κ��ļ�����</td></tr>"
		  else
		       totalPut=rs.recordcount
	        if currentpage<1 then
   		       currentpage=1
   	        end if
   	           if (currentpage-1)*MaxPerPage>totalput then
   		       if (totalPut mod MaxPerPage)=0 then
     		      currentpage= totalPut \ MaxPerPage
	  	       else
	      	      currentpage= totalPut \ MaxPerPage + 1
   		       end if
   	           end if
		if currentPage=1 then
	    Num=0
       	do while not rs.eof
		response.Write("<tr class=bgcolor_1><td><a href=Down.asp?id="&rs("id")&" target=_blank title="&htmlencode(rs("FileTitle"))&">"&getTopic(rs("FileTitle"),22)&"</a></td><td>"&rs("UserName")&"</td><td>"&myfilesize(rs("FileSize"))&"</td><td>"&rs("FileUpTime")&"</td><td>"&rs("IP")&"</td>")
		response.Write("<td align=center><a href=Admin.asp?action=delfile&id="&rs("id")&" onClick=""return confirm('��ȷ��Ҫִ��ɾ������?')"">ɾ��</a></td></tr>")
       	Num=Num+1
	   	if Num>=MaxPerPage then exit do
	   	rs.movenext
	    loop
		
 	else
     	if (currentPage-1)*MaxPerPage<totalPut then
       	   	rs.move  (currentPage-1)*MaxPerPage
       		dim bookmark
           	bookmark=rs.bookmark
        Num=0
       	do while not rs.eof
		response.Write("<tr class=bgcolor_1><td><a href=Down.asp?id="&rs("id")&" target=_blank title="&htmlencode(rs("FileTitle"))&">"&getTopic(rs("FileTitle"),22)&"</a></td><td>"&rs("UserName")&"</td><td>"&myfilesize(rs("FileSize"))&"</td><td>"&rs("FileUpTime")&"</td><td>"&rs("IP")&"</td>")
		response.Write("<td align=center><a href=Admin.asp?action=delfile&id="&rs("id")&" onClick=""return confirm('��ȷ��Ҫִ��ɾ������?')"">ɾ��</a></td></tr>")
       	Num=Num+1
	   	if Num>=MaxPerPage then exit do
	   	rs.movenext
	    loop
          
       	else
	        currentPage=1
        Num=0
       	do while not rs.eof
		response.Write("<tr class=bgcolor_1><td><a href=Down.asp?id="&rs("id")&" target=_blank title="&htmlencode(rs("FileTitle"))&">"&getTopic(rs("FileTitle"),22)&"</a></td><td>"&rs("UserName")&"</td><td>"&myfilesize(rs("FileSize"))&"</td><td>"&rs("FileUpTime")&"</td><td>"&rs("IP")&"</td>")
		response.Write("<td align=center><a href=Admin.asp?action=delfile&id="&rs("id")&" onClick=""return confirm('��ȷ��Ҫִ��ɾ������?')"">ɾ��</a></td></tr>")
       	Num=Num+1
	   	if Num>=MaxPerPage then exit do
	   	rs.movenext
	    loop
          	
	    end if
	end if
	response.Write("<tr><td colspan=6>"&showpage(strFileName,totalput,MaxPerPage,CurrentPage,true,true,"���ļ�")&"</td></tr>")
end if
rs.close
set rs=nothing
response.Write("</table>")
End Sub
%>
<%Call Bottom()%>