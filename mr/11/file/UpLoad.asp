<!--#include file="Conn.asp"-->
<!--#include file="Function.asp"-->
<!--#include file="inc.asp"-->
<!--#include file="UpLoadClass.asp"-->
<%If Session("UserName")="" Then
Response.Write("<script language=JavaScript>alert('����û�е�½��');document.location.href='Index.asp';</script>")
Response.End
End if
%>
<%Server.ScriptTimeOut=99999999%>
<%
Dim action
dim GroupInfo
GroupInfo=GetGroupInfo(Session("GroupID"))
action = trim(request.QueryString("action"))
SELECT Case action
	Case "SaveNew"
		Call SaveNew()
	Case else
		Call add()
End SELECT

Sub SaveNew()

If ChkPost=false Then
Response.Write("<script language=JavaScript>alert('�벻Ҫ���ⲿ�ύ��Ϣ��');document.location.href='Index.asp';</script>")
Response.End
End If

if GroupInfo(1)="" or GroupInfo(2)="" then
Response.Write("<script language=JavaScript>alert('���������ϴ��ļ�����Ҫ�ϴ��������Ա��ϵ��');document.location.href='Index.asp';</script>")
Response.End
End If
dim rs,sql
dim FileTitle,FileAbout,GroupID,ToUserName
dim upload,FileUrl,filesize,IP

set upload=New UpLoadClass   '�����ϴ�����
upload.open()

FileTitle=Checkstr(trim(upload.form("FileTitle")))
FileAbout=Checkstr(rtrim(upload.form("FileAbout")))
GroupID=Checkstr(trim(upload.form("GroupID")))
ToUserName=Checkstr(trim(upload.form("ToUserName")))
ToUserName=replace(ToUserName,"��",",")
if FileTitle = "" then
set upload=nothing
Response.Write("<script language=JavaScript>alert('�ļ����Ʋ���Ϊ�գ�');document.location.href='javascript:window.history.go(-1);';</script>")
exit sub
end if
if upload.Error=3 Then
set upload=nothing
Response.Write("<script language=JavaScript>alert('�ϴ�ʧ�ܣ��ļ�̫С��û��ѡ���ļ���');document.location.href='javascript:window.history.go(-1);';</script>")
exit sub 
end if
if upload.Error=1 Then
set upload=nothing
Response.Write("<script language=JavaScript>alert('�ϴ�ʧ�ܣ��ļ����������ϴ��Ĵ�С��');document.location.href='javascript:window.history.go(-1);';</script>")
exit sub 
end if
if upload.Error=2 Then
set upload=nothing
Response.Write("<script language=JavaScript>alert('�ϴ�ʧ�ܣ��������ϴ��������ļ���');document.location.href='javascript:window.history.go(-1);';</script>")
exit sub 
end if

FileUrl=upload.Form("FileUrl")
filesize=upload.filesize

if upload.Error<>0 OR FileUrl="" or FileTitle="" or filesize="" or filesize=<0 then
set upload=nothing 
Response.Write("<script language=JavaScript>alert('�ļ��ϴ�ʧ�ܣ���ȷ���Ѿ���ȷѡ���ϴ��ļ���');document.location.href='javascript:window.history.go(-1);';</script>")
exit sub 
end if
set upload=nothing ''ɾ���˶��� 
If Request.ServerVariables("HTTP_X_FORWARDED_FOR")<>"" then 
IP=Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
Else
IP=Request.ServerVariables("REMOTE_ADDR")
End If
set rs=server.CreateObject("adodb.recordset")
sql="select * from [tb_Files]"
If Not IsObject(Conn) Then ConnectionDatabase()
rs.open sql,conn,1,3
rs.addnew
rs("UserName")=Session("UserName")
rs("FileTitle")=FileTitle
rs("FileAbout")=FileAbout
rs("ToUserName")=ToUserName
rs("GroupID")=GroupID
rs("FileUrl")=FileUrl
rs("filesize")=filesize
rs("IP")=IP
rs.update
rs.close
set rs=nothing
Conn.Execute("Update tb_Users set UpFiles=UpFiles+1 Where UserName='"&Session("UserName")&"'")
CloseDatabase()
Response.Write("<script language=JavaScript>alert('�ļ��ϴ��ɹ����ļ���С��"&myfilesize(filesize)&" ');document.location.href='Index.asp';</script>")
End Sub

Sub add()
dim rs
dim styeid
Call Top("�ļ��ϴ�","UpLoad.asp")
%>
<table width=785 align="center">
  <tbody>
    <tr>
      <td style="font-size: 12px; width: 200px; height: 21px; text-align: center" 
    align=left rowspan=2 valign="top"><table class="table">
          <tbody>
            <tr>
              <td width="200" 
          colspan=3 bgcolor="#3399ff" style="font-weight: bold; width: 200px; color: white; height: 25px; text-align: center; text-decoration: none">�û���½</td>
            </tr>
            <tr>
              <td style="width: 200px; height: 145px; text-align: center" 
          colspan=3 rowspan=2><table style="width: 200px; height: 126px; text-align: left">
                  <tbody>
                    <tr>
                      <td style="width: 188px" colspan=3 rowspan=3><%call login()%></td>
                    </tr>
                    <tr></tr>
                    <tr></tr>
                  </tbody>
                </table></td>
            </tr>
            <tr></tr>
          </tbody>
        </table></td>
      <td style="width:100%; height: 21px" align=center valign="top" colspan=2 rowspan=2><table style="width: 100%; height: 182px">
          <tbody>
            <tr>
              <td colspan=3 bgcolor="#3399ff" style="font-weight: bold; color: white; height: 25px; text-align: center; text-decoration: none">�ϴ��ļ�</td>
            </tr>
            <tr>
              <td class="table" colSpan=3 rowSpan=3><form name="form" method="post" action="?action=SaveNew" enctype="multipart/form-data" onSubmit="javascript:Submit.disabled=true;Submit.value='���Ժ�...';">
                  <table border="0" align="center" cellpadding="2" cellspacing="1" width="100%">
                    <tr>
                      <td width="100" align="right">�ļ����ƣ�</td>
                      <td colspan="3"><input name="FileTitle" type="text" id="FileTitle" size="25" maxlength="60"></td>
                    </tr>
                    <tr>
                      <td align="right">�ļ�·����</td>
                      <td colspan="3"><input name="FileUrl" type="file" id="FileUrl" size="25">
                        <%if GroupInfo(1)="" or GroupInfo(2)="" Then
								  response.write "<br>���������ϴ��ļ�����Ҫ�ϴ��������Ա��ϵ��"
								  else
								  %>
                        <BR>
                        ��ֻ���ϴ�<font color="#FF0000"><strong><%=GroupInfo(2)%>KB</strong></font>�ڵ�<strong><%=replace(GroupInfo(1),"|",",")%></strong>�ļ�
                        <%end if%></td>
                    </tr>
                    <tr>
                      <td align="right" width="100">�ļ�˵����</td>
                      <td colspan="3"><textarea name="FileAbout" cols="60" rows="5" id="FileAbout"></textarea></td>
                    </tr>
                    <tr>
                      <td align="right">�����ļ���<br />
                        <script language="javascript">
								    function openScript(url, width, height)
								  {
	                               var Win = window.open(url,"openScript",'width=' + width + ',height=' + height + ',resizable=0,scrollbars=yes,menubar=no,status=no' );
                                  }
                                   </script>
                        <a href="javascript:openScript('UserList.asp',550,500)"><font color=red>�鿴�û��б�</font></a></td>
                      <td width="183"><%If Not IsObject(Conn) Then ConnectionDatabase()
								  set rs=Conn.execute("Select * from tb_Groups")
								   response.Write("<select name=""GroupID"" size=""5"" multiple=""multiple"" id=""GroupID"" style=""width:200px"">")
								   response.Write("<option value=""0"">�����û�</option>")
								   do until rs.eof
                                   response.Write("<option value="""&Rs("GroupID")&""">"&Rs("GroupName")&"</option>")
                                   rs.movenext
                                   loop
								   response.Write("</select>")
                                   set rs=nothing
								  %></td>
                      <td width="85" align="right">�����û���</td>
                      <td width="168"><input name="ToUserName" type="text" id="ToUserName" size="20" /></td>
                    </tr>
                    <tr>
                      <td colspan="4" align="center" height="30"><input type="submit" name="Submit" value="�ϴ�">
                        &nbsp;&nbsp;
                        <input type="reset" name="Submit2" value="����"></td>
                    </tr>
                    <tr>
                      <td colspan="4" align="left" style="word-break: break-all;FONT-SIZE: 9pt; COLOR: #552c55; LINE-HEIGHT: 120%; FONT-FAMILY: ����"><hr size="0" color="#eff7fe" />
                        <strong>˵��</strong>���ڹ����ļ�ѡ�ѡ����Ӧ������԰��ļ������������û�����סCtrl����ѡ�����飬<br/>
                        ���蹲��������û�ѡ�������û������ɣ����������ĳ���û������û����ڸ����û�����д<br/>
                        ���û����û������ɣ�����û�����ʹ�á�,���ָ�������д�����ļ�����ֻ������Լ���</td>
                    </tr>
                  </table>
                </form></td>
            </tr>
            <tr></tr>
            <tr></tr>
            <tr>
              <td class="titletd" colspan=3>&nbsp;</td>
            </tr>
          </tbody>
        </table></td>
    </tr>
    <tr></tr>
  </tbody>
</table>
<%Call Bottom()%>
<%
End Sub
%>
