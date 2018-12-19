<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
<style type="text/css">
<!--
.style1 {color: #f2ab5b}
.style2 {
	color: #f37b54;
	font-weight: bold;
	font-size: 14pt;
}
.style3 {font-weight: bold}
-->
</style>
<!--#include file="include/conn.asp" -->
<!--#include file="include/include.asp" -->
<%
if request("action")="add" then
	if trim(request("pinglun"))="" then
		response.Write("<script>alert('请详细填写！');history.back();</script>")
		response.End()
	end if
	sql="select * from [pinglun]"
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,3,3
	rs.addnew
	rs("shangpinid")=request("Id")
	rs("shijian")=now()
	rs("pinglun")=request("pinglun")
	rs("mingcheng")=request("mingcheng")
	rs.update
	rs.close
	set rs=nothing
	response.Write("<script>alert('评论发表成功！');window.location.href='lookpro.asp?id="&request("id")&"';</script>")
end if
%>
<table width="792" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td colspan="3"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="792" height="165" background="images/index_r1_c1.jpg"><!--#include file="top.asp" --></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td width="197"><!--#include file="left.asp" --></td>
    <td width="590" valign="top">
<table border="0" cellpadding="0" cellspacing="0" width="590">
<%
sql="select * from [shangpin] where id="&filter_Str(request("id"))&""	'filter_Str 函数的功能是检测是否有非法字符,代码在  include/include.asp
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
%>
  <tr>
    <td colspan="3"><img name="index_7_r1_c1" src="images/spxx.jpg" width="590" height="49" border="0" alt=""></td>
  </tr>
  <tr>
    <td colspan="3"><img name="index_7_r2_c1" src="images/index_7_r2_c1.jpg" width="590" height="9" border="0" alt=""></td>
  </tr>
  <tr>
    <td width="16" height="643" background="images/index_7_r3_c1.jpg">&nbsp;</td>
    <td width="565" valign="top">
<table width="94%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="173" rowspan="6"><div align="center"><img src="upfile/<%=rs("tupian")%>" width="110" height="130" border="0"></div></td>
    <td width="26" height="16">&nbsp;</td>
    <td width="323"><font color="EF9C3E">【商品名称】<%=rs("mingcheng")%></font></td>
    <td width="323" rowspan="2">【商品简介】<%=rs("jianjie")%></td>
  </tr>
  <tr>
    <td height="16">&nbsp;</td>
    <td><font color="910800">【市场价】<%=rs("shichang")%></font></td>
    </tr>
  <tr>
    <td height="16">&nbsp;</td>
    <td><font color="DD4679">【会员价】<%=rs("huiyuan")%></font></td>
    <td>【上架日期】<%=rs("riqi")%></td>
  </tr>
  <tr>
    <td height="16">&nbsp;</td>
    <td><a href="lookpro.asp?id=<%=rs("id")%>" target="_blank">【查看信息</a>】</td>
    <td>【商品型号】<%=rs("xinghao")%></td>
  </tr>
  <tr>
    <td height="16">&nbsp;</td>
    <td>【<a href="gouwu.asp?ProdId=<%=rs("id")%>">购买商品</a>】</td>
    <td>【商品等级】<%if rs("dengji")="2" then response.Write("精品") else response.Write("普通")%></td>
  </tr>
  <tr>
    <td height="16">&nbsp;</td>
    <td><font color="13589B">【浏览次数】<%=rs("cishu")%></font></td>
    <td>【商品数量】<%=rs("shuliang")%></td>
  </tr>
</table>
	<table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td height="30">&nbsp;</td>
      </tr>
    </table>
	<table width="96%"  border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td width="31%"><span class="style2">商品的详细说明：</span></td>
        <td width="69%">&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><%=HTMLEncode(rs("shuoming"))%></td>
      </tr>
    </table>
	<table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td height="30">&nbsp;</td>
      </tr>
    </table>	
	<table width="96%"  border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td width="31%"><span class="style2">商品的备注说明：</span></td>
        <td width="69%">&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><%=HTMLEncode(rs("beizhu"))%></td>
      </tr>
    </table>
	<table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td height="30">&nbsp;</td>
      </tr>
    </table>
	<table width="96%"  border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td><span class="style2">用户评论：</span>&nbsp;&nbsp;&nbsp;&nbsp;<a href="lookcomment.asp?id=<%=rs("id")%>" target="_blank">查看用户评论</a></td>
      </tr>
      <tr>
        <td>
		</td>
      </tr>
    </table>
	<table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td height="30">&nbsp;</td>
      </tr>
    </table>
	<table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td height="172">
<form action="lookpro.asp" method="post">
  <div align="center">
  <textarea name="pinglun" cols="60" rows="12"></textarea>
  <input type="hidden" name="id" value="<%=rs("id")%>">
  <input type="submit" value="评论">
  <input type="hidden" name="action" value="add">
  <input type="hidden" name="mingcheng" value="<%=rs("mingcheng")%>">
  </div>
</form>		</td>
      </tr>
    </table></td>
    <td width="10" background="images/index_7_r3_c3.jpg">&nbsp;</td>
  </tr>
  <tr>
    <td height="7" colspan="3"><img name="index_7_r4_c1" src="images/index_7_r4_c1.jpg" width="590" height="7" border="0" alt=""></td>
  </tr>
<%
rs("cishu")=rs("cishu")+1
rs.update
rs.close
set rs=nothing
%>
</table></td>
    <td width="8"><img name="index_r2_c3" src="images/index_r2_c3.jpg" width="5" height="753" border="0" alt=""></td>
  </tr>
  <tr>
    <td colspan="3"><!--#include file="foot.asp" --></td>
  </tr>
</table>