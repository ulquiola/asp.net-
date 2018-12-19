<!--#include file="Conn.asp"-->
<!--#include file="Function.asp"-->
<!--#include file="inc.asp"-->
<%
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1  
Response.cachecontrol = "no-cache" 
Dim Rs,Sql
If Session("UserName")="" Then
Response.Write("<script language=JavaScript>alert('您还没有登陆！');document.location.href='Index.asp';</script>")
Response.End
End if

If Request.Form("Action")="Edit" Then
If ChkPost=false Then
Response.Write("<script language=JavaScript>alert('请不要重外部提交信息！');document.location.href='Index.asp';</script>")
Response.End
End If
Dim UserPwd,UserSex,Company,Department,Tel,Mobile
UserPwd=Trim(Request.Form("UserPwd"))
UserSex=Checkstr(Trim(Request.Form("UserSex")))
Company=Checkstr(Trim(Request.Form("Company")))
Department=Checkstr(Trim(Request.Form("Department")))
Tel=Checkstr(Trim(Request.Form("Tel")))
Mobile=Checkstr(Trim(Request.Form("Mobile")))

if Tel<>"" Then
if IsValidTel(Tel)=False then
Response.Write("<script language=JavaScript>alert('电话号码格式错误！');document.location.href='javascript:window.history.go(-1);';</script>")
Response.End
End If
end if

if Mobile<>"" then
If isInteger(Mobile)=False Then
Response.Write("<script language=JavaScript>alert('手机号码必须是数字！');document.location.href='javascript:window.history.go(-1);';</script>")
Response.End
End If
end if

Sql="Select * From tb_Users Where UserName='"&Session("UserName")&"'"
Set Rs=server.CreateObject("adodb.recordset")
If Not IsObject(Conn) Then ConnectionDatabase()
Rs.open Sql,Conn,1,3
If Rs.Eof Then
Rs.close
Set Rs=Nothing
CloseDatabase()
Response.Write("<script language=JavaScript>alert('用户不存在！');document.location.href='javascript:window.history.go(-1);';</script>")
Else

Rs("UserPwd")=UserPwd
Rs("UserSex")=CBool(UserSex)
Rs("Company")=Company
Rs("Department")=Department
Rs("Tel")=Tel
Rs("Mobile")=Mobile
Rs.UpDate
Rs.close
Set Rs=nothing
CloseDatabase()
Response.Write("<script language=JavaScript>alert('资料修改成功！');document.location.href='Edit.asp';</script>")
End if
Response.End
End If
%>

<%Call Top("修改资料","Edit.asp")%>
<table width=785 align="center">
  <tbody>
  <tr>
    <td style="font-size: 12px; width: 200px; height: 21px; text-align: center" 
    align=left rowspan=2 valign="top">
      <table class="table">
        <tbody>
        <tr>
          <td width="200" 
          colspan=3 bgcolor="#3399ff" style="font-weight: bold; width: 200px; color: white; height: 25px; text-align: center; text-decoration: none">用户登陆</td>
        </tr>
        <tr>
          <td style="width: 200px; height: 145px; text-align: center" 
          colspan=3 rowspan=2>
            <table style="width: 200px; height: 126px; text-align: left">
              <tbody>
              <tr>
                <td style="width: 188px" colspan=3 rowspan=3><%call login()%></td></tr>
              <tr></tr>
              <tr></tr></tbody></table></td></tr>
        <tr></tr></tbody></table></td>
    <td style="width:100%; height: 21px" align=center valign="top" colspan=2 rowspan=2>
      <table style="width: 100%; height: 182px">
        <tbody>
        <tr>
          <td colspan=3  bgcolor="#3399ff" style="font-weight: bold; color: white; height: 25px; text-align: center; text-decoration: none">修改资料</td>
        </tr>
        <tr>
          <td class="table" colspan=3 rowspan=3 valign="top"><form id="form1" name="form1" method="post" action="edit.asp">
		  <%If Not IsObject(Conn) Then ConnectionDatabase()
		  Set Rs=Conn.execute("Select * from tb_Users Where UserName='"&Session("UserName")&"'")
		  if rs.eof then
		     set Rs=nothing
			 CloseDatabase()
			 Response.Write("<script language=JavaScript>alert('用户不存在，请重新登陆！');document.location.href='javascript:window.history.go(-1);';</script>")
			 Response.End
		  End if
		  %>
            <input name="Action" type="hidden" id="Action" value="Edit" />
            <table width="100%" border="0" cellspacing="3" cellpadding="0">
              <tr>
                <td width="21%" align="right">用户名：</td>
                <td width="79%"><%=rs("UserName")%></td>
              </tr>
              <tr>
                <td  align="right">密　码：</td>
                <td><input name="UserPwd" type="text" id="UserPwd" maxlength="20" /> 
                  不修改请不要填写。 </td>
              </tr>
              <tr>
                <td  align="right">性　别：</td>
                <td><input type="radio" name="UserSex" value="1" <%if rs("UserSex") then response.Write("checked")%> />
                  男
                    <input type="radio" name="UserSex" value="0" <%if not rs("UserSex") then response.Write("checked")%> />
                    女</td>
              </tr>
              <tr>
                <td align="right">单　位：</td>
                <td><input name="Company" type="text" id="Company" value="<%=rs("Company")%>" maxlength="48" /></td>
              </tr>
              <tr>
                <td align="right">部　门：</td>
                <td><input name="Department" type="text" id="Department" value="<%=rs("Department")%>" maxlength="48" /></td>
              </tr>
              <tr>
                <td align="right">手　机：</td>
                <td><input name="Mobile" type="text" id="Mobile" value="<%=rs("Mobile")%>" maxlength="20" /></td>
              </tr>
              <tr>
                <td align="right">电　话：</td>
                <td><input name="Tel" type="text" id="Tel" value="<%=rs("Tel")%>" maxlength="20" /></td>
              </tr>
              <tr>
                <td align="right">注　册：</td>
                <td><%=rs("UserRegTime")%></td>
              </tr>
              <tr>
                <td align="right">登　陆：</td>
                <td><%=rs("UserLastTime")%></td>
              </tr>
              <tr>
                <td align="right">登陆IP：</td>
                <td><%=rs("UserLastIp")%></td>
              </tr>
              <tr>
                <td colspan="2" align="center"><input type="submit" name="Submit" value="修改" /></td>
                </tr>
            </table>
                    </form>
          </td>
        </tr>
        <tr></tr>
        <tr></tr>
        <tr>
          <td class="titletd" colspan=3>&nbsp;</td>
        </tr></tbody></table></td></tr>
  <tr></tr></tbody></table>
<%Call Bottom()%>