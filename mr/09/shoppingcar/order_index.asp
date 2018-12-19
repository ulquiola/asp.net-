<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/conn.asp"--><!--数据库连接文件-->
<%
OrderID=request.QueryString("OrderID")						'获取订单编号
set rs=server.CreateObject("adodb.recordset")				'创建记录集
sql="select * from tb_order where OrderID='"&OrderID&"'"	 '查询数据
rs.open sql,conn,1,3									 	'打开记录集
  	if not rs.eof  then 								 
  		truename=rs("truename")								'为变量赋值
  		address=rs("address")
  		postcode=rs("postcode")								'为变量赋值
  		tel=rs("tel")
   		pay=rs("pay")
  		carry=rs("carry")									'为变量赋值
  		orderdate=rs("orderdate")
  		bz=rs("bz")
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<style type="text/css">
<!--
.STYLE4 {color: #FF0000}
.STYLE6 {color: #FF0000; font-size: 9pt; }
-->
</style>
</head>
<style type="text/css">
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-bottom: 0px;
}
</style>
<!--#include file="top.asp"-->
<body>
<table width="800" height="528" border="0" align="center" cellpadding="0" cellspacing="0" background="images/new17.gif">
  <tr>
    <td height="86" colspan="3">&nbsp;</td>
  </tr>
  <tr>
    <td width="46" height="289">&nbsp;</td>
    <td width="712" valign="top"><table width="693" height="70" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td width="693" height="70">
		
		<table width="381" height="300" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="376" valign="middle" background="../images/0big.gif"><table width="94%" height="273" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td height="231" valign="top"></div>
          <table width="357" height="230" border="0" align="center" cellpadding="0" cellspacing="0" background="../images/hbig.gif">
          <tr>
            <td width="28%" height="17" valign="bottom"><div align="center" class="STYLE2">真实姓名</div></td>
            <td width="72%" valign="bottom"><input name="truename" type="text" id="truename" size="25" value="<%=rs("truename")%>"></td>
          </tr>
          <tr>
            <td height="19"><div align="center" class="STYLE2">联系地址</div></td>
            <td height="19"><input name="address" type="text" id="address" size="25" value="<%=rs("address")%>" /></td>
          </tr>
          <tr>
            <td height="25"><div align="center" class="STYLE2">邮政编码</div></td>
            <td height="25"><input name="postcode" type="text" id="postcode" size="25" value="<%=rs("postcode")%>" /></td>
          </tr>
          <tr>
            <td height="28"><div align="center" class="STYLE2">联系电话</div></td>
            <td height="28"><input name="tel" type="text" id="tel" size="25" value="<%=rs("tel")%>" /></td>
          </tr>
          <tr>
            <td height="24"><div align="center" class="STYLE2">付款方式</div></td>
            <td height="24"><select name="pay" id="pay">
			<option value="银行付款"<%if(instr(rs("pay"),"银行付款")>0)then%>selected<%end if%>>银行付款</option>
			<option value="邮政付款"<%if(instr(rs("pay"),"邮政付款")>0)then%>selected<%end if%>>邮政付款</option>
			<option value="现金支付"<%if(instr(rs("pay"),"现金支付")>0)then%>selected<%end if%>>现金支付</option>
            </select></td>
          </tr>
          <tr>
            <td height="15"><div align="center" class="STYLE2">运送方式</div></td>
            <td height="15"><select name="carry" id="carry">
			<option value="普通邮寄"<%if(instr(rs("carry"),"普通邮寄")>0)then%>selected<%end if%>>普通邮寄</option>
			<option value="特快专递"<%if(instr(rs("carry"),"特快专递")>0)then%>selected<%end if%>>特快专递</option>
			<option value="EMS专递方式"<%if(instr(rs("carry"),"EMS专递方式")>0)then%>selected<%end if%>>EMS专递方式</option>
           </select></td>
          </tr>
          <tr>
            <td height="8"><div align="center" class="STYLE2">订单日期</div></td>
            <td height="8"><input name="orderdate" type="text" id="orderdate" size="25" value="<%=rs("orderDate")%>" /></td>
          </tr>
          <tr>
            <td height="8"><div align="center" class="STYLE2">订单是否执行</div></td>
            <td height="8">
			<%if rs("enforce")=0 then%>
			<span class="STYLE6">未执行</span>
			<%end if%>
			<%if rs("enforce")=1 then%>
			<span class="STYLE6">已执行</span>
			<%end if%>			</td>
          </tr>
          <tr>
            <td height="26"><div align="center">备&nbsp;&nbsp;&nbsp;注</div></td>
            <td height="26"><textarea name="bz" cols="30" rows="3" id="bz"><%=rs("bz")%>
</textarea></td>
          </tr>
          <tr>
            <td height="29" colspan="2" valign="bottom"><div align="center"><a href="#" onclick="javascript:window.close()">【关闭窗口】</a></div></td>
            </tr>
        </table></td>
      </tr>
      
    </table></td>
  </tr>
</table>
		
		</td>
      </tr>
    </table></td>
    <td width="42">&nbsp;</td>
  </tr>
  <tr>
    <td height="13" colspan="3">&nbsp;</td>
  </tr>
  <tr>
    <td height="129" colspan="3">&nbsp;</td>
  </tr>
</table>
 <% else %>
  <script language="javascript">
  alert("订单信息不存在");				//弹出提示对话框
  window.close();						//关闭窗口
  </script>
<%end if %>
</body>
</html>
