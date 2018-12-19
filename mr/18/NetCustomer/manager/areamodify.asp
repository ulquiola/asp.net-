<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Connections/conn.asp" -->
<%
if request.Form("HqyNO")<>"" then
	HqyNO=request.Form("HqyNO")
	UP="Update DB_WorkerInfo set workqy="& HqyNO &" where ywynumber='"&session("ywynumber")&"'"
	conn.execute(UP)
	%>
	  <script language="javascript">
	  alert("工作区域修改成功！");
	  opener.location.reload();
	  window.close();
	  </script>
<%
end if
if Request.QueryString("CityID")="" then
session("SelectID_P")=request.QueryString("ProvinceID")
end if
SelectID_C=Request.QueryString("CityID")
if request.QueryString("ProvinceID")="" and Request.QueryString("CityID")="" then
session("ywynumber")=request("ywynumber")
end if
Set rs_province = Server.CreateObject("ADODB.Recordset")
sql_P = "SELECT *  FROM DB_Area  WHERE TypeID=1 or TypeID=4 or TypeID=5"
rs_province.open sql_P,conn,1,3
If (Session("ywynumber") <> "") Then 
  NO = Session("ywynumber")
End If
Set rs_ywy = Server.CreateObject("ADODB.Recordset")
sql_ywy = "SELECT ywynumber, name, workqy FROM DB_WorkerInfo WHERE ywynumber ='"&NO&"'"
rs_ywy.open sql_ywy,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>管理员登录--地域选择！</title>
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
<script language="javascript">
function ChangeItem_P(){
if(form1.Province.value!=0){
var ProvinceID=form1.Province.value;
window.location.href="areamodify.asp?ProvinceID="+ProvinceID;}
}
</script>
<script language="javascript">
function ChangeItem_C(){
if(form1.city.value!=0){
var CityID=form1.city.value;
window.location.href="areamodify.asp?CityID="+CityID;}
else{
form1.town.value=0;}
}
</script>
<body>
<script language="javascript">
function mycheck(){
  if(form1.town.value==0){
     if(form1.city.value==0){
        if(form1.Province.value==0){
	       alert("请先输入地域信息！");
		   window.close();
	    }
		 else
   {form1.HqyNO.value=form1.Province.value;}
     }
	 else
   {form1.HqyNO.value=form1.city.value;}
   }
   else
   {form1.HqyNO.value=form1.town.value;}
   form1.submit();
}
</script>
<table width="400" height="242" border="0" cellpadding="-2" cellspacing="-2"
 background="../Images/add_manager.gif">
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
            <form ACTION="areamodify.asp" name="form1" method="POST">
            
            <table width="88%" height="142" border="1" cellpadding="-2"  cellspacing="-2"
			 bordercolor="#669999" bordercolordark="#FFFFFF">
              <tr>
                <td width="32%" height="27" bgcolor="#809EA4"><div align="center" class="style1">
				省级名称： </div></td>
                <td width="68%">
				<%
				if session("SelectID_P")<>"" then
				sql="select* from DB_area where TypeID=2 and father="&session("SelectID_P")
				sql_T="select* from DB_area where TypeID=3 and father="&session("SelectID_P")
				
				else
				sql="select* from DB_area where TypeID=2 and father="&rs_province("ID")
				sql_T="select* from DB_area where TypeID=3 and father="&rs_province("ID")
				end if
				set rs_city=conn.execute(sql)
				set rs_Town=conn.execute(sql_T)
				%><div align="left">&nbsp;
				  <select name="Province" id="select4" onChange="ChangeItem_P()">
  <%
 rs_province.movefirst  
While (NOT rs_province.EOF)
%>
  <option value="<%=rs_province("ID")%>" <% if rs_province("ID")=cint(session("SelectID_P")) then %>
  selected<% end if%>><%=(rs_province.Fields.Item("Name").Value)%></option>
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
                  </div></td>
                </tr>
              <tr>
                <td height="27" bgcolor="#809EA4"><div align="center" class="style1">市/县名称：
                  </div></td>
                <td><div align="left">&nbsp;
                <select name="city" id="select4" onChange="ChangeItem_C()">
<%
While (NOT rs_city.EOF)
%>
  <option value="<%=rs_city("ID")%>" <% if rs_city("ID")=cint(SelectID_C) then %>
   selected <%end if%>><%=(rs_city.Fields.Item("Name").Value)%></option>
  <%
  rs_city.MoveNext()
Wend
If (rs_city.CursorType > 0) Then
  rs_city.MoveFirst
Else
  rs_city.Requery
End If
%>
<option value="0">---</option>
</select>
</div></td></tr>
              <tr>
                <td height="27" bgcolor="#809EA4"><div align="center" class="style1">
				区/镇/乡名称：</div></td>
                <td><div align="left">&nbsp;
<select name="town" id="select5"><%
				if SelectID_C<>"" then
				sql_T="select* from DB_area where father="&SelectID_C
				set rs_town=conn.execute(sql_T)
				else
				if rs_Town.eof then
				if not rs_city.eof then
				sql_T="select* from DB_area where father="&rs_city("ID")
				set rs_town=conn.execute(sql_T)
				else
				sql_T="select* from DB_area where father=-1"
				set rs_town=conn.execute(sql_T)
				end if
				end if
				end if
				%>			 
<% While (NOT rs_town.EOF)
%><option value="<%=rs_town("ID")%>" <% if rs_town("ID")=SelectID_C then %> selected 
<%end if%>><%=rs_town("Name")%></option>
  <%
  rs_town.MoveNext()
Wend%>
<option value="0">---</option>
</select>
<%
If (rs_town.CursorType > 0) Then
  rs_town.MoveFirst
Else
  rs_town.Requery
End If%>

</div></td>
              </tr>
              <tr>
                <td height="33" colspan="2"><div align="center">
                  <input name="Button" type="button" class="Style_button" value="确定"
				   onClick="mycheck()">
                  <input name="Button" type="button" class="Style_button" value="关闭"
				   onClick="javascrip:opener.location.reload();window.close()">
				  <input name="HqyNO" type="hidden" id="HqyNO">
                </div></td>
              </tr>
              </table>
			</form>
            </div></td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td></td>
        <td>&nbsp;</td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
</html>
