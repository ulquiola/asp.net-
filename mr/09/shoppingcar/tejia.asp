<!--#include file="conn/conn.asp"--><%'包含数据库连接文件%>
<%
Set rs_sale=server.CreateObject("ADODB.RecordSet")                  	  '创建记录集
sql="select top 3 * from tb_goods where sale=1 order by INTime desc"	  '查询数据
rs_sale.open sql,conn,1,3												  '打开记录集
%>
<%
If(rs_sale.Eof)Then															'判断是否有记录信息
	Response.Write("暂无商品信息!")											'弹出提示信息
Else
%>
<style type="text/css">
<!--
.STYLE2 {
	font-size: 9pt;
	color: #000000;
}
.STYLE4 {font-size: 9pt}
body {
	margin-bottom: 0px;
}
-->
</style>
<br><br><br>
<table width="500" border="0" cellpadding="0" cellspacing="0" align="center">
<tr>
<%
for i=1 to rs_sale.recordcount		'统计记录总数
if i mod 3=0 then					'分栏显示商品信息
%>
<td height="185" valign="top"><table width="82%"  border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="95" align="center"><img src="images/goods/<%=rs_sale("picture")%>" width="150" height="100" border="1"></td>
                    </tr>
                    <tr>
                      <td height="18" align="center" class="STYLE2"><%=rs_sale("goodsname")%></td>
                    </tr>
                    <tr>
<td height="18" align="center" style="text-decoration:line-through;color:#FF0000" class="STYLE3">原价：<%=rs_sale("price")%> </td>
                    </tr>
                    <tr>
                      <td height="18" align="center" class="STYLE2">现价：<%=rs_sale("nowprice")%></td>
                    </tr>
                    <tr>
                      <td height="29" align="center"><a href="#" onClick="javascript:window.open('chakan.asp?id=<%=rs_sale("ID")%>&type=new','','width=415,height=299,toolbar=no,location=no,status=no,menubar=no')"><img src="images/new8.gif" width="72" height="26" border="0" /></a> &nbsp;<img src="images/mai.gif" width="57" height="24" onclick="window.location.href='cart_add.asp?goodsID=<%=rs_sale("ID")%>'" /></td>
                    </tr>
    </table></td>
</tr> 
<%else%>
<td height="185" valign="top"><table width="82%"  border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td height="95" align="center"><img src="images/goods/<%=rs_sale("picture")%>" width="150" height="100" border="1"></td>
                      </tr>
                      <tr>
                        <td height="18" align="center" class="STYLE2"><%=rs_sale("goodsname")%></td>
                      </tr>
                      <tr>
                        <td height="18" align="center" class="STYLE3" style="text-decoration:line-through;color:#FF0000">原价：<%=rs_sale("price")%> </td>
                      </tr>
                      <tr>
                        <td height="18" align="center" class="STYLE2">现价：<%=rs_sale("nowprice")%></td>
                      </tr>
                      <tr>
                        <td height="29" align="center"><a href="#" onClick="javascript:window.open('chakan.asp?id=<%=rs_sale("ID")%>&type=new','','width=415,height=299,toolbar=no,location=no,status=no,menubar=no')"><img src="images/new8.gif" width="72" height="26" border="0" /></a> &nbsp;<img src="images/mai.gif" width="57" height="24" onclick="window.location.href='cart_add.asp?goodsID=<%=rs_sale("ID")%>'" /></td>
                      </tr>
    </table></td>
<%
end if
rs_sale.movenext		'向下移动记录指针
next
%> 
</table>
<%end if%>