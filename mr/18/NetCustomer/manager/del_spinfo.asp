<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Connections/conn.asp" -->
<%
'创建显商品信息的记录集
If (Request.QueryString("spnumber") <> "") Then 
  session("spno") = Request.QueryString("spnumber")
End If
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT * FROM DB_SPInfo WHERE spnumber = '" + Replace(session("spno"), "'", "''") + "'"
rs.open sql,conn,1,3
'删除商品信息
if request.Form("spname")<>"" then
	Del="Delete From DB_SPInfo WHERE spnumber = '" + Replace(session("spno"), "'", "''") + "'"
	conn.execute(Del) %>
	  <script language="javascript">
	  alert("数据删除成功！");
	  opener.location.href="manage_sp.asp";
	  window.close();
	  </script>
<%response.End()
end if%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>商品信息！</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style.css" rel="stylesheet">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	background-image: url(../Images/bg.gif);
}
.style1 {color: #FFFFFF}
.style2 {color: #a2bcc5}
-->
</style></head>

<body>
<table width="595" height="242" border="0" cellpadding="-2" cellspacing="-2"
 background="../Images/manage.gif">
  <tr>
    <td valign="top"><table width="595" height="427" cellpadding="-2" cellspacing="-2">
      <tr>
        <td width="10" height="112">&nbsp;</td>
        <td width="577">&nbsp;</td>
        <td width="10"><span class="style2"></span></td>
      </tr>
      <tr>
        <td height="258">&nbsp;</td>
        <td valign="top"><div align="center">
          <form ACTION="del_spinfo.asp" METHOD="POST" name="form1">
            <table border="1" align="center" cellpadding="-2" cellspacing="-2"
			 bordercolor="#669999" bordercolordark="#FFFFFF">
              <tr valign="baseline">
                <td width="77" height="30" align="right" nowrap bgcolor="#809EA4">
				<div align="center" class="style1">商品编号：</div></td>
                <td width="102" valign="middle">
                  <div align="left">
                    &nbsp;<input name="spnumber" type="text" class="Sytle_text"
					 value="<%=(rs.Fields.Item("spnumber").Value)%>" size="32" readonly="yes">
                  </div></td>
                <td width="88" bgcolor="#809EA4"><div align="center" class="style1">
				商品简称：</div></td>
                <td width="119" valign="middle"><div align="left">&nbsp;
				<input name="spsymbol" type="text" class="Sytle_auto"
				 value="<%=rs("spsymbol")%>" size="15">
                </div></td>
              </tr>
              <tr valign="baseline">
                <td height="30" align="right" nowrap bgcolor="#809EA4">
				<div align="center" class="style1">商品名称：</div></td>
                <td colspan="3">
                  <div align="left">
                    &nbsp;<input name="spname" type="text" class="Style_upload"
					 value="<%=(rs.Fields.Item("spname").Value)%>" size="32">
                  </div></td>
              </tr>
              <tr valign="baseline">
                <td height="30" align="right" nowrap bgcolor="#809EA4">
				<div align="center" class="style1">包　　装：</div></td>
                <td valign="middle">
                  <div align="left">
                    &nbsp;<input name="pack" type="text" class="Sytle_text"
					 value="<%=(rs.Fields.Item("pack").Value)%>" size="32">
                  </div></td>
                <td bgcolor="#809EA4"><div align="center" class="style1">类　　型：</div></td>
                <td valign="middle"><div align="left">&nbsp;<input name="type"
				 type="text" class="Sytle_auto" value="<%=rs("type")%>" size="15">
                </div></td>
              </tr>
              <tr valign="baseline">
                <td height="30" align="right" nowrap bgcolor="#809EA4"><div align="center" class="style1">单　　价：</div></td>
                <td valign="middle">
                  <div align="left">
                    &nbsp;<input name="price" type="text" class="Sytle_text" value="<%=(rs.Fields.Item("price").Value)%>" size="32">              
                      </div></td>
                <td bgcolor="#809EA4"><div align="center" class="style1">口味类型：</div></td>
                <td valign="middle"><div align="left">&nbsp;<input name="kwtype" type="text" class="Sytle_auto" value="<%=(rs.Fields.Item("kwtype").Value)%>" size="15">
                </div></td>
              </tr>
              <tr valign="baseline">
                <td height="30" align="right" nowrap bgcolor="#809EA4"><div align="center" class="style1">生产厂商：</div></td>
                <td colspan="3" valign="middle">
                  <div align="left">&nbsp;<input name="sccs" type="text" class="Style_upload" id="sccs" value="<%=(rs.Fields.Item("sccs").Value)%>" size="32">
</div></td>
              </tr>
              <tr valign="baseline">
                <td height="30" align="right" nowrap bgcolor="#809EA4"><div align="center" class="style1">备　　注：</div></td>
                <td colspan="3" valign="middle">
                  <div align="left">
                    &nbsp;<input name="memo" type="text" class="Style_upload" value="<%=(rs.Fields.Item("memo").Value)%>" size="32">
                  </div></td>
              </tr>
              <tr valign="middle">
                <td height="38" colspan="4" align="right" nowrap>
                  <div align="center">
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
        <td height="21">&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
</html>
