<!--#include file="Conn.asp"-->
<!--#include file="Function.asp"-->
<!--#include file="inc.asp"-->
<%
Dim Rs,Sql,id
if request.QueryString("action")="down" then
dim fileurl
If Session("UserName")="" Then
Response.Write("<script language=JavaScript>alert('����û�е�½��ֻ�е�½�˲������أ�');document.location.href='index.asp';</script>")
Response.End
End if
id=Trim(Request.QueryString("id"))
if id="" Then
Response.write"������ʧ��"
Elseif not isnumeric(id) Then
response.Write("�������ʹ���")
else
Sql="Select * From tb_Files where id="&clng(id)
set Rs=server.CreateObject("adodb.recordset")
If Not IsObject(Conn) Then ConnectionDatabase()
Rs.open sql,conn,1,3
if not rs.eof then
if rs("UserName")<>Session("UserName") and instr(","+Rs("GroupID")+",","0")=0 and instr(","+Rs("GroupID")+",",Session("GroupID"))=0 and instr(","+Lcase(Rs("ToUserName"))+",",Lcase(Session("UserName")))=0 and Not Session("UserAdmin") Then
rs.close
set rs=nothing
CloseDatabase()
response.write("����Ȩ���ظ��ļ���")
else
Rs("FileDowns")=Rs("FileDowns")+1
Rs.update
fileurl=Rs("FileUrl")
rs.close
set rs=nothing
CloseDatabase()
response.redirect "UploadFile/"&fileurl
end if
else
rs.close
set rs=nothing
CloseDatabase()
response.write("�ļ������ڣ�")
end if
end if
response.End
end if
%>
<%Call Top("�ļ�����","Down.asp")%>
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
              <td colspan=3  bgcolor="#3399ff" style="font-weight: bold; color: white; height: 25px; text-align: center; text-decoration: none">�ļ�����</td>
            </tr>
            <tr>
              <td class="table" colspan=3 rowspan=3><%
					  id=Trim(Request.QueryString("id"))
					  if id="" Then
					  Response.write"������ʧ��"
					  Elseif not isnumeric(id) Then
					  response.Write("�������ʹ���")
					  Else
					  Sql="Select * From tb_Files where id="&clng(id)
					  set Rs=server.CreateObject("adodb.recordset")
					  If Not IsObject(Conn) Then ConnectionDatabase()
					  Rs.open sql,conn,1,1
					  if rs.eof then
					  response.Write("���ļ������ڣ�")
					  elseif rs("UserName")<>Session("UserName") and instr(","+Rs("GroupID")+",","0")=0 and instr(","+Rs("GroupID")+",",Session("GroupID"))=0 and instr(","+Lcase(Rs("ToUserName"))+",",Lcase(Session("UserName")))=0 and Not Session("UserAdmin") Then
					  response.write("����Ȩ���ظ��ļ���")
					  else
					  response.write "<div align=center><b>"&HTMLEncode(Rs("FileTitle"))&"</b><br>�ϴ�ʱ�䣺"&rs("FileUpTime")&"&nbsp;&nbsp;���ش�����"&rs("FileDowns")&"&nbsp;&nbsp;�ļ���С����"&myfilesize(rs("FileSize"))&"</div><hr size=0 color=""#cae2f8"">"
					  response.write HTMLEncode(Rs("FileAbout"))
					  response.write("<br><a href=down.asp?action=down&id="&id&"><img src=""images/download.gif"" width=""22"" height=""22"" border=""0"" /></a>")
				      End If
				      Rs.close
				      Set Rs=nothing
					  End if
				   %></td>
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
