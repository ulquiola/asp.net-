<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Connections/conn.asp" -->
<%
'查询商品信息数据表中最大的商品编号
Set rs_Max = Server.CreateObject("ADODB.Recordset")
sql_max = "select Max(spnumber) as MaxNO  from DB_Spinfo"
rs_Max.open sql_max,conn,1,3
'保存商品信息
if request.Form("spname")<>"" then
	spsymbol=request.Form("spsymbol")
	spname=request.Form("spname")
	pack=request.Form("pack")
	ctype=request.Form("type")
	price=request.Form("price")
	kwtype=request.Form("kwtype")
	sccs=request.Form("sccs")
	memo=request.Form("memo")
	Ins="Insert into DB_SPInfo (spnumber,spsymbol,spname,pack,type,price,kwtype,sccs,memo)"&_
	" values('"&session("autoNO")&"','"&spsymbol&"','"&spname&"','"&pack&"','"&ctype&"',"&_
	price&",'"&kwtype&"','"&sccs&"','"&memo&"')"
	conn.execute(Ins)%>
	  <script language="javascript">
	  alert("数据添加成功！");
	  opener.location.reload();
	  window.close();
	  </script>
<%end if %>
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
<script language="javascript">
function mycheck(){
if (form1.spname.value=="")
{alert("请输入商品名称！");form1.spname.focus();return;}
if(form1.sccs.value=="")
{alert("请输入生产厂商！");form1.sccs.focus();return;}
if (isNaN(form1.price.value))　//判断用户输入的单价是否为数值型
{alert("您输入的单价不正确！");form1.price.focus();return;}
form1.submit();
}
</script>
<% 
Dim no,valno,mymonth,cmonth,cday
no=trim(rs_Max("MaxNO"))
'生成五位的编号
select case len(int(Right(no,5)+1))
	case 1
		cno="0000"+Cstr(int(Right(no,5)+1))
	case 2
		cno="000"+Cstr(int(Right(no,5)+1))
	case 3
		cno="00"+Cstr(int(Right(no,5)+1))
	case 4
		cno="0"+Cstr(int(Right(no,5)+1))
	case 5
		cno=Cstr(int(Right(no,5)+1))
	case Else
		cno="00001"
end select
intno="SP"+cno '组成新的数据编号
session("autoNO")=intno  '将生成的数据编号赋值给会话变量
%>
<table width="595" height="242" border="0" cellpadding="-2" cellspacing="-2" background="../Images/manage.gif">
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
          <form method="POST" action="add_spinfo.asp" name="form1">
            <table border="1" align="center" cellpadding="-2" cellspacing="-2" bordercolor="#669999" bordercolordark="#FFFFFF">
              <tr>
                <td width="77" height="30" align="right" nowrap bgcolor="#809EA4"><div align="center" class="style1">商品编号：</div></td>
                <td width="102">
                  <div align="left">
                    &nbsp;<input name="spnumber" type="text" class="Sytle_text" value="<%=session("autoNO")%>" size="32" readonly="yes">
                  </div></td>
                <td width="88" bgcolor="#809EA4"><div align="center" class="style1">商品简称：</div></td>
                <td width="119"><div align="left">&nbsp;<input name="spsymbol" type="text" class="Sytle_auto" value="" size="15">
                </div></td>
              </tr>
              <tr>
                <td height="30" align="right" nowrap bgcolor="#809EA4"><div align="center" class="style1">商品名称：</div></td>
                <td colspan="3">                  <div align="left">
                    &nbsp;<input name="spname" type="text" class="Style_upload" value="" size="32">
                  </div></td>
              </tr>
              <tr>
                <td height="30" align="right" nowrap bgcolor="#809EA4"><div align="center" class="style1">包　　装：</div></td>
                <td>
                  <div align="left">
                    &nbsp;<input name="pack" type="text" class="Sytle_text" value="" size="32">
                  </div></td>
                <td bgcolor="#809EA4"><div align="center" class="style1">类　　型：</div></td>
                <td><div align="left">&nbsp;<input name="type" type="text" class="Sytle_auto" value="" size="15">
                </div></td>
              </tr>
              <tr>
                <td height="30" align="right" nowrap bgcolor="#809EA4"><div align="center" class="style1">单　　价：</div></td>
                <td>
                  <div align="left">
                    &nbsp;<input name="price" type="text" class="Sytle_auto" value="" size="9">
                    (元)              
                   </div></td>
                <td bgcolor="#809EA4"><div align="center" class="style1">口味类型：</div></td>
                <td><div align="left">&nbsp;<input name="kwtype" type="text" class="Sytle_auto" value="" size="15">
                </div></td>
              </tr>
              <tr>
                <td height="30" align="right" nowrap bgcolor="#809EA4"><div align="center" class="style1">生产厂商：</div></td>
                <td colspan="3">
                  <div align="left">&nbsp;<input name="sccs" type="text" class="Style_upload" id="sccs" value="" size="32">
</div></td>
              </tr>
              <tr>
                <td height="30" align="right" nowrap bgcolor="#809EA4"><div align="center" class="style1">备　　注：</div></td>
                <td colspan="3">
                  <div align="left">
                    &nbsp;<input name="memo" type="text" class="Style_upload" value="" size="32">
                  </div></td>
              </tr>
              <tr valign="middle">
                <td height="38" colspan="4" align="right" nowrap>
                  <div align="center">
                    <input name="Button" type="button" class="Style_button" value="保存"
					 onClick="mycheck()">                
                    <input name="Reset" type="reset" class="Style_button" value="重置">
                    <input name="Button" type="button" class="Style_button" value="关闭"
					 onClick="javascrip:opener.location.reload();window.close()">
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
