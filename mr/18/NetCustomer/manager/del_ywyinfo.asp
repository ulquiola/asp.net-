<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Connections/conn.asp" -->
<!--#include file="../../NetCustomer/F_Display.asp" -->
<%''创建显示业务员信息的记录集
If (Request.QueryString("ywynumber") <> "") Then 
  session("NO") = Request.QueryString("ywynumber")
End If
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT * FROM DB_WorkerInfo WHERE ywynumber = '"&session("NO")& "'"
rs.open sql,conn,1,3%>
<%'删除业务员信息
if request.Form("name")<>"" then
	Del="Delete from DB_WorkerInfo where ywynumber = '"&session("NO")& "'"
	conn.execute(Del)%>
	<script language="javascript">
	alert("数据删除成功！");
	opener.location.href="manage_ywy.asp";//重新链接网页
	window.close();
	</script>
<% end if%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>业务员信息！</title>
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
<table width="595" height="242" border="0" cellpadding="-2" cellspacing="-2" background="../Images/manage.gif">
  <tr>
    <td valign="top"><table width="595" cellpadding="-2" cellspacing="-2">
      <tr>
        <td width="10" height="112">&nbsp;</td>
        <td width="577">&nbsp;</td>
        <td width="10"><span class="style2"></span></td>
      </tr>
      <tr>
        <td height="258">&nbsp;</td>
        <td valign="top"><div align="center">
          <form ACTION="del_ywyinfo.asp" METHOD="POST" name="form1">
            <table border="1" align="center" cellpadding="-2" cellspacing="-2" bordercolor="#669999" bordercolordark="#FFFFFF">
              <tr valign="middle">
                <td width="72" align="right" nowrap bgcolor="#809EA4"><div align="center" class="style1">编　　号：</div></td>
                <td width="105">                    <div align="left">
                  <input name="ywynumber" type="text" class="Sytle_text" id="ywynumber" value="<%=(rs.Fields.Item("ywynumber").Value)%>" size="32" readonly="yes">
                </div></td>
                <td width="71" bgcolor="#809EA4"><div align="center" class="style1">姓　　名：</div></td>
                <td width="96"><div align="left">
                  <input name="name" type="text" class="Sytle_text" value="<%=(rs.Fields.Item("name").Value)%>" size="32">
                </div></td>
                <td width="68" bgcolor="#809EA4"><div align="center" class="style1">性　　别：</div></td>
                <td width="97"><div align="left">
                  <select name="sex" id="select3">
                    <option value="男" <%If (Not isNull((rs.Fields.Item("sex").Value))) Then If ("男" = CStr((rs.Fields.Item("sex").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>男</option>
                    <option value="女" <%If (Not isNull((rs.Fields.Item("sex").Value))) Then If ("女" = CStr((rs.Fields.Item("sex").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>女</option>
                  </select>
                </div></td>
              </tr>
              <tr valign="middle">
                <td align="right" nowrap bgcolor="#809EA4"><div align="center" class="style1">民　　族：</div></td>
                <td><div align="left">
                  <input name="folk" type="text" class="Sytle_text" value="<%=(rs.Fields.Item("folk").Value)%>" size="32">
                </div></td>
                <td bgcolor="#809EA4"><div align="center" class="style1">生　　日：</div></td>
                <td><div align="left">
                  <input name="birthday" type="text" class="Sytle_text" value="<%=(rs.Fields.Item("birthday").Value)%>" size="32">
                </div></td>
                <td bgcolor="#809EA4"><div align="center" class="style1">入厂日期：</div></td>
                <td><div align="left">
                  <input name="rcdate" type="text" class="Sytle_text" value="<%=(rs.Fields.Item("rcdate").Value)%>" size="32">
                </div></td>
              </tr>
              <tr valign="middle">
                <td align="right" nowrap bgcolor="#809EA4"><div align="center" class="style1">学　　历：</div></td>
                <td><div align="left">
                  <select name="xl" id="select4">
                    <option value="博士及以上" selected <%If (Not isNull((rs.Fields.Item("xl").Value))) Then If ("博士及以上" = CStr((rs.Fields.Item("xl").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>博士及以上</option>
                    <option value="硕士" <%If (Not isNull((rs.Fields.Item("xl").Value))) Then If ("硕士" = CStr((rs.Fields.Item("xl").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>硕士</option>
                    <option value="研究生" <%If (Not isNull((rs.Fields.Item("xl").Value))) Then If ("研究生" = CStr((rs.Fields.Item("xl").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>研究生</option>
                    <option value="本科" <%If (Not isNull((rs.Fields.Item("xl").Value))) Then If ("本科" = CStr((rs.Fields.Item("xl").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>本科</option>
                    <option value="专科" <%If (Not isNull((rs.Fields.Item("xl").Value))) Then If ("专科" = CStr((rs.Fields.Item("xl").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>专科</option>
                    <option value="高中及中专" <%If (Not isNull((rs.Fields.Item("xl").Value))) Then If ("高中及中专" = CStr((rs.Fields.Item("xl").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>高中及中专</option>
                    <option value="初中及以下" <%If (Not isNull((rs.Fields.Item("xl").Value))) Then If ("初中及以下" = CStr((rs.Fields.Item("xl").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>初中及以下</option>
                  </select>
</div></td>
                <td bgcolor="#809EA4"><div align="center" class="style1">身份证号：</div></td>
                <td colspan="3"><div align="left">
                  <input name="sfzNO" type="text" class="Sytle_auto" value="<%=(rs.Fields.Item("sfzNO").Value)%>" size="32">
                </div></td>
              </tr>
              <tr valign="middle">
                <td align="right" nowrap bgcolor="#809EA4"><div align="center" class="style1">专　　业：</div></td>
                <td><div align="left">
                  <select name="zy" id="select">
                    <option value="计算机" selected <%If (Not isNull((rs.Fields.Item("zy").Value))) Then If ("计算机" = CStr((rs.Fields.Item("zy").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>计算机</option>
                    <option value="食品" <%If (Not isNull((rs.Fields.Item("zy").Value))) Then If ("食品" = CStr((rs.Fields.Item("zy").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>食品</option>
                    <option value="机械制造" <%If (Not isNull((rs.Fields.Item("zy").Value))) Then If ("机械制造" = CStr((rs.Fields.Item("zy").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>机械制造</option>
                    <option value="医生" <%If (Not isNull((rs.Fields.Item("zy").Value))) Then If ("医生" = CStr((rs.Fields.Item("zy").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>医生</option>
                    <option value="护士" <%If (Not isNull((rs.Fields.Item("zy").Value))) Then If ("护士" = CStr((rs.Fields.Item("zy").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>护士</option>
                    <option value="教师" <%If (Not isNull((rs.Fields.Item("zy").Value))) Then If ("教师" = CStr((rs.Fields.Item("zy").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>教师</option>
                    <option value="其他" <%If (Not isNull((rs.Fields.Item("zy").Value))) Then If ("其他" = CStr((rs.Fields.Item("zy").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>其他</option>
                  </select>
                </div></td>
                <td bgcolor="#809EA4"><div align="center" class="style1">联系电话：</div></td>
                <td colspan="3"><div align="left">
                  <input name="tel" type="text" class="Sytle_auto" value="<%=(rs.Fields.Item("tel").Value)%>" size="32">
                </div></td>
                </tr>
              <tr valign="middle">
                <td align="right" nowrap bgcolor="#809EA4"><div align="center" class="style1">工作区域：</div></td>
                <td colspan="5"><div align="left">                  <table width="98%"  cellspacing="-2" cellpadding="-2">
                    <tr>
                      <td width="86%">
						<% qyNO=rs("workqy")
						call Display(qyNO)  '显示工作区域%>
					  </td>
                      <td width="14%"><input name="workqy" type="hidden" id="workqy4" value="<%=rs("workqy")%>"></td>
                    </tr>
                  </table>
</div>                  <div align="left"></div></td>
                </tr>
              <tr valign="middle">
                <td align="right" nowrap bgcolor="#809EA4"><div align="center" class="style1">地　　址：</div></td>
                <td colspan="5">
                  <div align="left">
                    <input name="address" type="text" class="Style_upload" value="<%=(rs.Fields.Item("address").Value)%>" size="32">                  
                  </div></td>
              </tr>
              <tr valign="middle">
                <td align="right" nowrap bgcolor="#809EA4"><div align="center" class="style1">Email:</div></td>
                <td colspan="5">
                  <div align="left">
                    <input name="email" type="text" class="Sytle_auto" value="<%=(rs.Fields.Item("email").Value)%>" size="32">                
                  </div></td>
              </tr>
              <tr valign="middle">
                <td align="right" nowrap bgcolor="#809EA4"><div align="center" class="style1">备　　注：</div></td>
                <td colspan="5">
                  <div align="left">
                    <input name="memo" type="text" class="Style_upload" value="<%=(rs.Fields.Item("memo").Value)%>" size="32">                
                  </div></td>
              </tr>
              <tr valign="middle">
                <td height="35" colspan="6" nowrap><div align="center">
                    <input name="Submit" type="submit" class="Style_button" value="删除">                
                    <input name="Submit2" type="button" class="Style_button" value="关闭" onClick="JScript:window.close()">
                </div></td>
                </tr>
            </table>
          </form>
          <p>&nbsp;</p>
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
