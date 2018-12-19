<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../../Connections/conn.asp" -->
<% set conn=server.CreateObject("ADODB.Connection")
conn.open("Driver={Microsoft Access Driver (*.mdb)};PWD=111;DBQ=" & server.MapPath("/NetCustomer/Customer.mdb"))
session("SelectID_P")=request.QueryString("ProvinceID")
SelectID_C=Request.QueryString("CityID")
%><%
Dim rs_province
Dim rs_province_numRows

Set rs_province = Server.CreateObject("ADODB.Recordset")
rs_province.ActiveConnection = MM_conn_STRING
rs_province.Source = "SELECT *  FROM DB_Area  WHERE TypeID=1 or TypeID=4 or TypeID=5"
rs_province.CursorType = 0
rs_province.CursorLocation = 2
rs_province.LockType = 1
rs_province.Open()

rs_province_numRows = 0
%>
<%
Dim rs_province_select__MMColParam
rs_province_select__MMColParam = "1"
If (Request.Form("province") <> "") Then 
  rs_province_select__MMColParam = Request.Form("province")
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>管理员登录--删除地域信息！</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style.css" rel="stylesheet">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	background-image:  url(../Images/bg.gif);
}
.style1 {color: #FFFFFF}
.style2 {color: #a2bcc5}
-->
</style></head>

<body>
<script language="javascript">
function ChangeItem_P(){
var ProvinceID=form1.Province.value;
window.location.href="del_area_T.asp?ProvinceID="+ProvinceID;
}
</script>
<script language="javascript">
function ChangeItem_C(){
var CityID=form1.city.value;
var ProvinceID=form1.HProvince.value;
window.location.href="del_area_T.asp?CityID="+CityID;
}
</script>
<table width="400" height="242" border="0" cellpadding="-2" cellspacing="-2" background="../Images/add_manager.gif">
  <tr>
    <td valign="top"><table width="400" height="271" cellpadding="-2" cellspacing="-2">
      <tr>
        <td width="11" height="85">&nbsp;</td>
        <td width="373">&nbsp;</td>
        <td width="14"><span class="style2"></span></td>
      </tr>
      <tr>
        <td height="144">&nbsp;</td>
        <td valign="top">
  
          <div align="center">
            <form name="form1" method="post">
            
            <table width="80%" height="142" border="1" cellpadding="-2"  cellspacing="-2" bordercolor="#669999" bordercolordark="#FFFFFF">
              <tr>
                <td width="32%" height="27" bgcolor="#809EA4"><div align="center" class="style1">省级名称： </div></td>
                <td width="68%">
				<%
				if session("SelectID_P")<>"" then
				sql="select* from DB_area where father="&session("SelectID_P")
				set rs_city=conn.execute(sql)
				else
				session("SelectID_P")=1
				sql="select* from DB_area where father=1"
				set rs_city=conn.execute(sql)
				end if
				%>                  <div align="left">&nbsp;
				  <select name="Province" id="select4" onChange="ChangeItem_P()">
  <%
 rs_province.movefirst  
While (NOT rs_province.EOF)
%>
  <option value="<%=(rs_province.Fields.Item("ID").Value)%>" <% if rs_province("ID")=cint(session("SelectID_P")) then %>selected<% end if%>><%=(rs_province.Fields.Item("Name").Value)%></option>
  <%
  rs_province.MoveNext()
Wend
If (rs_province.CursorType > 0) Then
  rs_province.MoveFirst
Else
  rs_province.Requery
End If
%>
</select>
<!--隐藏表单*****************************************-->
                  <input name="HProvince" type="hidden" id="HProvince" 
				  <% if selectID<>""then %> value="<%=session("SelectID_P") %>"
				  <% else %>
					value="1"
				  <% end if %>
				 ><!--********************-->
                  </div></td>
                </tr>
              <tr>
                <td height="27" bgcolor="#809EA4"><div align="center" class="style1">市/县名称：
                  </div></td>
                <td height="34"><div align="left">&nbsp;
                <select name="city" id="select4" onChange="ChangeItem_C()">
  <%
While (NOT rs_city.EOF)
%>
  <option value="<%=(rs_city.Fields.Item("ID").Value)%>" <% if rs_city("ID")=session("SelectID_P") then %> selected <%end if%>><%=(rs_city.Fields.Item("Name").Value)%></option>
  <%
  rs_city.MoveNext()
Wend
If (rs_city.CursorType > 0) Then
  rs_city.MoveFirst
Else
  rs_city.Requery
End If
IF rs_city.EOF then %>
<option value="0">该省级名称没有下属地区！</option>
<% End If %>
</select>
</div></td>
              </tr>
              <tr>
                <td height="27" bgcolor="#809EA4"><div align="center" class="style1">区/镇/乡名称：</div></td>
                <td><div align="left">&nbsp;
                  				<%
				if SelectID_C<>"" then
				sql_T="select* from DB_area where father="&SelectID_C
				set rs_town=conn.execute(sql_T)
				else
				if not rs_city.eof then
				sql_T="select* from DB_area where father="&rs_city("ID")
				set rs_town=conn.execute(sql_T)
				end if
				end if
				%>                  <div align="left">&nbsp;
                <select name="town" id="select5">
  <%
While (NOT rs_town.EOF)
%>
  <option value="<%=(rs_town.Fields.Item("ID").Value)%>" <% if rs_town("ID")=SelectID_C then %> selected <%end if%>><%=(rs_town.Fields.Item("Name").Value)%></option>
  <%
  rs_town.MoveNext()
Wend
If (rs_town.CursorType > 0) Then
  rs_town.MoveFirst
Else
  rs_town.Requery
End If
IF rs_city.EOF then %>
<option value="0">该区/镇名称没有下属地区！</option>
<% End If %>
</select>
</div></td>
              </tr>
              <tr>
                <td height="33" colspan="2"><div align="center">
                  <input name="Button" type="button" class="Style_button" value="保存" onClick="mycheck()">
                  <input name="Submit2" type="reset" class="Style_button" value="重置">
                  <input name="Button" type="button" class="Style_button" value="关闭" onClick="javascrip:opener.location.reload();window.close()">
</div></td>
              </tr>
              </table>
            </form>
            </div></td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><form name="form2" method="post">
          <input name="HProvince" type="hidden" id="HProvince">
		</form></td>
        <td>&nbsp;</td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
</html>
<%
rs_town.Close()
Set rs_town = Nothing
%>
<%
rs_province.Close()
Set rs_province = Nothing
%>

<%
rs_city.close()
set rs_city=nothing
%>
