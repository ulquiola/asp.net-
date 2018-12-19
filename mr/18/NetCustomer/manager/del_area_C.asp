<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Connections/conn.asp" -->  <!--包含数据库连接文件-->
<% 
Set rs_province = Server.CreateObject("ADODB.Recordset")
sql_P = "SELECT *  FROM DB_Area  WHERE TypeID=1 or TypeID=4 or TypeID=5"
rs_province.open sql_P,conn,1,3
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
function ChangeItem(){
var ProvinceID=form1.Province.value;
window.location.href="del_area_C.asp?ProvinceID="+ProvinceID; 
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
             <form name="form1" method="post" action="del_area_C_OK.asp">
            
            <table width="80%" height="88" border="1" cellpadding="-2"  cellspacing="-2" bordercolor="#669999" bordercolordark="#FFFFFF">
              <tr>
                <td width="32%" height="27" bgcolor="#809EA4"><div align="center" class="style1">省级名称： </div></td>
                <td width="68%">
				<%
				SelectID=request.QueryString("ProvinceID")
				if SelectID<>"" then
				sql="select* from DB_area where father="&SelectID
				set rs_city=conn.execute(sql)
				else
				sql="select* from DB_area where father="&rs_province("ID")
				set rs_city=conn.execute(sql)
				end if
				%>
                  <div align="left">&nbsp;
                   <select name="Province" id="select" onChange="ChangeItem()">
                   <% rs_province.movefirst  '移动到第一条记录
					While (NOT rs_province.EOF)%>
                      <option value="<%=rs_province("ID")%>" <% if rs_province("ID")=cint(SelectID) then %>
					  selected<% end if%>><%=rs_province("Name")%></option>
                      <% rs_province.MoveNext()   '向下移到一条记录
					Wend%>
                   </select>
                  </div></td>
                </tr>
              <tr>
                <td height="27" bgcolor="#809EA4"><div align="center" class="style1">市/县名称：
                  </div></td>
                <td height="34"><div align="left">&nbsp;
<select name="city" id="select4">
  <% While (NOT rs_city.EOF)%>
  <option value="<%=rs_city("ID")%>" <% if rs_city("ID")=SelectID then %>
   selected <%end if%>><%=rs_city("Name")%></option>
  <% rs_city.MoveNext()
  Wend
  IF rs_city.EOF and rs_city.bof then    '当记录集为空时%>
 <option value="0">---</option>
  <% End If %>
</select>
</div></td>
              </tr>
              <tr>
                <td height="33" colspan="2"><div align="center">
                  <input name="Submit" type="submit" class="Style_button" value="删除">
                  <input name="Button" type="button" class="Style_button" value="关闭"
				   onClick="javascrip:opener.location.reload();window.close()">
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
