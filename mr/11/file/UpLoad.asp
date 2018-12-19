<!--#include file="Conn.asp"-->
<!--#include file="Function.asp"-->
<!--#include file="inc.asp"-->
<!--#include file="UpLoadClass.asp"-->
<%If Session("UserName")="" Then
Response.Write("<script language=JavaScript>alert('您还没有登陆！');document.location.href='Index.asp';</script>")
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
Response.Write("<script language=JavaScript>alert('请不要重外部提交信息！');document.location.href='Index.asp';</script>")
Response.End
End If

if GroupInfo(1)="" or GroupInfo(2)="" then
Response.Write("<script language=JavaScript>alert('您不允许上传文件，需要上传请与管理员联系！');document.location.href='Index.asp';</script>")
Response.End
End If
dim rs,sql
dim FileTitle,FileAbout,GroupID,ToUserName
dim upload,FileUrl,filesize,IP

set upload=New UpLoadClass   '建立上传对象
upload.open()

FileTitle=Checkstr(trim(upload.form("FileTitle")))
FileAbout=Checkstr(rtrim(upload.form("FileAbout")))
GroupID=Checkstr(trim(upload.form("GroupID")))
ToUserName=Checkstr(trim(upload.form("ToUserName")))
ToUserName=replace(ToUserName,"，",",")
if FileTitle = "" then
set upload=nothing
Response.Write("<script language=JavaScript>alert('文件名称不能为空！');document.location.href='javascript:window.history.go(-1);';</script>")
exit sub
end if
if upload.Error=3 Then
set upload=nothing
Response.Write("<script language=JavaScript>alert('上传失败，文件太小或没有选择文件！');document.location.href='javascript:window.history.go(-1);';</script>")
exit sub 
end if
if upload.Error=1 Then
set upload=nothing
Response.Write("<script language=JavaScript>alert('上传失败，文件超过您能上传的大小！');document.location.href='javascript:window.history.go(-1);';</script>")
exit sub 
end if
if upload.Error=2 Then
set upload=nothing
Response.Write("<script language=JavaScript>alert('上传失败，不允许上传此类型文件！');document.location.href='javascript:window.history.go(-1);';</script>")
exit sub 
end if

FileUrl=upload.Form("FileUrl")
filesize=upload.filesize

if upload.Error<>0 OR FileUrl="" or FileTitle="" or filesize="" or filesize=<0 then
set upload=nothing 
Response.Write("<script language=JavaScript>alert('文件上传失败！请确认已经正确选择上传文件！');document.location.href='javascript:window.history.go(-1);';</script>")
exit sub 
end if
set upload=nothing ''删除此对象 
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
Response.Write("<script language=JavaScript>alert('文件上传成功！文件大小："&myfilesize(filesize)&" ');document.location.href='Index.asp';</script>")
End Sub

Sub add()
dim rs
dim styeid
Call Top("文件上传","UpLoad.asp")
%>
<table width=785 align="center">
  <tbody>
    <tr>
      <td style="font-size: 12px; width: 200px; height: 21px; text-align: center" 
    align=left rowspan=2 valign="top"><table class="table">
          <tbody>
            <tr>
              <td width="200" 
          colspan=3 bgcolor="#3399ff" style="font-weight: bold; width: 200px; color: white; height: 25px; text-align: center; text-decoration: none">用户登陆</td>
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
              <td colspan=3 bgcolor="#3399ff" style="font-weight: bold; color: white; height: 25px; text-align: center; text-decoration: none">上传文件</td>
            </tr>
            <tr>
              <td class="table" colSpan=3 rowSpan=3><form name="form" method="post" action="?action=SaveNew" enctype="multipart/form-data" onSubmit="javascript:Submit.disabled=true;Submit.value='请稍后...';">
                  <table border="0" align="center" cellpadding="2" cellspacing="1" width="100%">
                    <tr>
                      <td width="100" align="right">文件名称：</td>
                      <td colspan="3"><input name="FileTitle" type="text" id="FileTitle" size="25" maxlength="60"></td>
                    </tr>
                    <tr>
                      <td align="right">文件路径：</td>
                      <td colspan="3"><input name="FileUrl" type="file" id="FileUrl" size="25">
                        <%if GroupInfo(1)="" or GroupInfo(2)="" Then
								  response.write "<br>您不允许上传文件，需要上传请与管理员联系。"
								  else
								  %>
                        <BR>
                        您只能上传<font color="#FF0000"><strong><%=GroupInfo(2)%>KB</strong></font>内的<strong><%=replace(GroupInfo(1),"|",",")%></strong>文件
                        <%end if%></td>
                    </tr>
                    <tr>
                      <td align="right" width="100">文件说明：</td>
                      <td colspan="3"><textarea name="FileAbout" cols="60" rows="5" id="FileAbout"></textarea></td>
                    </tr>
                    <tr>
                      <td align="right">共享文件：<br />
                        <script language="javascript">
								    function openScript(url, width, height)
								  {
	                               var Win = window.open(url,"openScript",'width=' + width + ',height=' + height + ',resizable=0,scrollbars=yes,menubar=no,status=no' );
                                  }
                                   </script>
                        <a href="javascript:openScript('UserList.asp',550,500)"><font color=red>查看用户列表</font></a></td>
                      <td width="183"><%If Not IsObject(Conn) Then ConnectionDatabase()
								  set rs=Conn.execute("Select * from tb_Groups")
								   response.Write("<select name=""GroupID"" size=""5"" multiple=""multiple"" id=""GroupID"" style=""width:200px"">")
								   response.Write("<option value=""0"">所有用户</option>")
								   do until rs.eof
                                   response.Write("<option value="""&Rs("GroupID")&""">"&Rs("GroupName")&"</option>")
                                   rs.movenext
                                   loop
								   response.Write("</select>")
                                   set rs=nothing
								  %></td>
                      <td width="85" align="right">附加用户：</td>
                      <td width="168"><input name="ToUserName" type="text" id="ToUserName" size="20" /></td>
                    </tr>
                    <tr>
                      <td colspan="4" align="center" height="30"><input type="submit" name="Submit" value="上传">
                        &nbsp;&nbsp;
                        <input type="reset" name="Submit2" value="重置"></td>
                    </tr>
                    <tr>
                      <td colspan="4" align="left" style="word-break: break-all;FONT-SIZE: 9pt; COLOR: #552c55; LINE-HEIGHT: 120%; FONT-FAMILY: 宋体"><hr size="0" color="#eff7fe" />
                        <strong>说明</strong>：在共享文件选项处选择相应的组可以把文件共享给该组的用户，按住Ctrl可以选择多个组，<br/>
                        如需共享给所有用户选择“所有用户”即可，单独共享给某个用户或多个用户则在附加用户处填写<br/>
                        该用户的用户名即可，多个用户必须使用“,”分隔。不填写共享文件处则只共享给自己。</td>
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
