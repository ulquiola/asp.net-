<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Connections/conn.asp" -->
<!--#include file="../../NetCustomer/F_Display.asp" -->
<%
'创建显客户信息的记录集
If (Request.QueryString("khnumber") <> "") Then 
  session("NO") = Request.QueryString("khnumber")
End If
Set rs_kh = Server.CreateObject("ADODB.Recordset")
sql_NO = "SELECT * FROM DB_KhInfo WHERE khnumber = '"&session("NO")& "'"
rs_kh.open sql_NO,conn,1,3
%>
<%'修改客户信息
if request.Form("khname") <>"" then
	khname=request.Form("khname")
	khsymbol=request.Form("khsymbol")
	khtype=request.Form("khtype")
	khgrade=request.Form("khgrade")
	sellsum=request.Form("sellsum")
	frdb=request.Form("frdb")
	postcode=request.Form("postcode")
	address=request.Form("address")
	tel=request.Form("tel")
	email=request.Form("email")
	fax=request.Form("fax")
	homepage=request.Form("homepage")
	khh=request.Form("khh")
	zh=request.Form("zh")
	memo=request.Form("memo")
	Up="Update DB_KhInfo set khname='"&khname&"',khsymbol='"&khsymbol&"',khtype='"&khtype&_
	"',khgrade='"&khgrade&"',sellsum="&sellsum&",frdb='"&frdb&"',postcode='"&postcode&_
	"',address='"&address&"',tel='"&tel&"',email='"&email&"',fax='"&fax&"',homepage='"&_
	homepage&"',khh='"&khh&"',zh='"&zh&"',memo='"&khname&"'WHERE khnumber = '"&session("NO")&"'"
	conn.execute(UP)
	%>
	  <script language="javascript">
	  alert("数据修改成功！");
	  opener.location.reload();
	  window.close();
	  </script>
<% end if %>	  
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>客户信息！</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style.css" rel="stylesheet">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	background-image: url(../Images/bg.gif);
}
.style2 {color: #a2bcc5}
.style4 {color: #FFFFFF}
-->
</style></head>

<body>
<script language="javascript">
function mycheck(){
if (form1.khname.value=="")
{alert("请输入客户名称！");form1.khname.focus();return;}
if(form1.frdb.value=="")
{alert("请输入法人代表！");form1.frdb.focus();return;}
if(form1.tel.value=="")
{alert("请输入电话号码！");form1.tel.focus();return;}
form1.submit();
}
</script>
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
          <form ACTION="modify_khinfo.asp" METHOD="POST" name="form1">
            <table width="539" border="1" align="center" bordercolor="#669999" bordercolordark="#FFFFFF" cellpadding="-2" cellspacing="-2">
              <tr valign="baseline">
                <td width="60" align="right" nowrap bgcolor="#809EA4"><div align="center" class="style4">客户编号：</div></td>
                <td width="103">
                  <div align="left">
                    &nbsp;
                    <input name="khnumber" type="text" class="Sytle_text" value="<%=rs_kh("khnumber")%>" size="32" readonly="yes">
                  </div></td>
                <td width="70" bgcolor="#809EA4"><span class="style4">客户名称：</span></td>
                <td colspan="3"><div align="left">
                  &nbsp;
                  <input name="khname" type="text" class="Sytle_auto" value="<%=rs_kh("khname")%>" size="32">
                </div></td>
              </tr>
              <tr valign="baseline">
                <td align="right" nowrap bgcolor="#809EA4"><div align="center" class="style4">客户简称：</div></td>
                <td valign="middle">
                  <div align="left">
                    &nbsp;
                    <input name="khsymbol" type="text" class="Sytle_text" value="<%=rs_kh("khsymbol")%>" size="32">
                  </div></td>
                <td valign="middle" bgcolor="#809EA4"><span class="style4">客户类型：</span></td>
                <td width="125" valign="middle"><div align="left">
                  &nbsp;
                  <input name="khtype" type="text" class="Sytle_text" value="<%=rs_kh("khtype")%>" size="32">
                </div></td>
                <td width="64" valign="middle" bgcolor="#809Ea4"><span class="style4">客户等级：</span></td>
                <td width="89" valign="middle"><div align="left">
                    &nbsp;
                    <select name="khgrade" id="select">
                      <option value="A" <%If (Not isNull((rs_kh.Fields.Item("khgrade").Value))) Then If ("A" = CStr((rs_kh.Fields.Item("khgrade").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>A</option>
                      <option value="B" <%If (Not isNull((rs_kh.Fields.Item("khgrade").Value))) Then If ("B" = CStr((rs_kh.Fields.Item("khgrade").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>B</option>
                      <option value="C" <%If (Not isNull((rs_kh.Fields.Item("khgrade").Value))) Then If ("C" = CStr((rs_kh.Fields.Item("khgrade").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>C</option>
                      <option value="D" <%If (Not isNull((rs_kh.Fields.Item("khgrade").Value))) Then If ("D" = CStr((rs_kh.Fields.Item("khgrade").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>D</option>
                    </select>
                    (级)</div></td>
              </tr>
              <tr valign="baseline">
                <td align="right" nowrap bgcolor="#809EA4"><div align="center" class="style4">销售数量：</div></td>
                <td valign="middle">
                  <div align="left">
                    &nbsp;
                    <input name="sellsum" type="text" class="Sytle_text" id="sellsum" value="<%=(rs_kh.Fields.Item("sellsum").Value)%>" size="32">
                  </div></td>
                <td valign="middle" bgcolor="#809EA4"><span class="style4">法人代表：</span></td>
                <td colspan="3" valign="middle"><div align="left">
                  &nbsp;
                  <input name="frdb" type="text" class="Sytle_text" value="<%=(rs_kh.Fields.Item("frdb").Value)%>" size="32">
                </div></td>
                </tr>
              <tr valign="baseline">
                <td align="right" nowrap bgcolor="#809EA4"><div align="center" class="style4">所属区域：</div></td>
                <td colspan="5" valign="middle">
                  <div align="left">
&nbsp;
				    <% qyNO=rs_kh("qyname")
				  call Display(qyNO)
				 %>
				    <input type="hidden" name="qyname" value="<%=rs_kh("qyname")%>">
                  </div></td>
              </tr>
              <tr valign="baseline">
                <td align="right" nowrap bgcolor="#809EA4"><div align="center" class="style4">邮政编码：</div></td>
                <td valign="middle"><div align="left">
                  &nbsp;
                  <input name="postcode" type="text" class="Sytle_text" value="<%=(rs_kh.Fields.Item("postcode").Value)%>" size="32">
                </div></td>
                <td valign="middle" bgcolor="#809EA4"><span class="style4">地　　址：</span></td>
                <td colspan="3" valign="middle"><div align="left">
                  &nbsp;
                  <input name="address" type="text" class="Sytle_auto" value="<%=(rs_kh.Fields.Item("address").Value)%>" size="32">
                </div></td>
                </tr>
              <tr valign="baseline">
                <td align="right" nowrap bgcolor="#809EA4"><div align="center" class="style4">电　　话：</div></td>
                <td valign="middle">
                  <div align="left">
                    &nbsp;
                    <input name="tel" type="text" class="Sytle_text" value="<%=(rs_kh.Fields.Item("tel").Value)%>" size="32">
                  </div></td>
                <td valign="middle" bgcolor="#809EA4"><span class="style4">Email：</span></td>
                <td colspan="3" valign="middle"><div align="left">
                  &nbsp;
                  <input name="email" type="text" class="Sytle_auto" value="<%=(rs_kh.Fields.Item("email").Value)%>" size="32">
                </div></td>
                </tr>
              <tr valign="baseline">
                <td align="right" nowrap bgcolor="#809EA4"><div align="center" class="style4">传　　真：</div></td>
                <td valign="middle">
                  <div align="left">
                    &nbsp;
                    <input name="fax" type="text" class="Sytle_text" value="<%=(rs_kh.Fields.Item("fax").Value)%>" size="32">
                  </div></td>
                <td valign="middle" bgcolor="#809EA4"><span class="style4">网　　址：</span></td>
                <td colspan="3" valign="middle"><div align="left">
                  &nbsp;
                  <input name="homepage" type="text" class="Sytle_auto" value="<%=(rs_kh.Fields.Item("homepage").Value)%>" size="32">
                </div></td>
                </tr>
              <tr valign="baseline">
                <td align="right" nowrap bgcolor="#809EA4"><div align="center" class="style4">开 户 行：</div></td>
                <td valign="middle">
                  <div align="left">
                    &nbsp;
                    <input name="khh" type="text" class="Sytle_text" value="<%=(rs_kh.Fields.Item("khh").Value)%>" size="32">
                  </div></td>
                <td valign="middle" bgcolor="#809EA4"><span class="style4">账　　号：</span></td>
                <td colspan="3" valign="middle"><div align="left">
                  &nbsp;
                  <input name="zh" type="text" class="Sytle_auto" value="<%=rs_kh("zh")%>" size="32">
                </div></td>
                </tr>
              <tr valign="baseline">
                <td align="right" nowrap bgcolor="#809EA4"><div align="center" class="style4">备　　注：</div></td>
                <td colspan="5" valign="middle">
                  <div align="left">
                    &nbsp;
                    <input name="memo" type="text" class="Style_subject" value="<%=rs_kh("memo")%>" size="32">
                  </div></td>
              </tr>
              <tr valign="baseline">
                <td colspan="6" align="center" valign="middle" nowrap>
				<input type="button" class="Style_button" value="保存" onclick="mycheck()">
				<input name="Reset" type="reset" class="Style_button" value="重置">
				<input name="Submit2" type="button" class="Style_button" value="关闭"
				 onClick="JScript:window.close()"></td>
                </tr>
            </table>
          </form>
          <p>&nbsp;</p>
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
