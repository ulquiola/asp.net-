<!--#include file="conn/conn.asp"-->
<%
Set rs_sale=server.CreateObject("ADODB.RecordSet")			'创建记录集
sql="select * from tb_goods  order by INTime desc"			'查询数据
rs_sale.open sql,conn,1,3									'打开记录集
%>
<html>
<head>
<title>购买商品</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<style type="text/css">
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-bottom: 0px;
}
</style>
<body>
<!--#include file="top.asp"-->
<table width="800" height="529" border="0" align="center" cellpadding="0" cellspacing="0" background="images/new19.gif">
  <tr>
    <td height="46" colspan="3">&nbsp;</td>
  </tr>
  <tr>
    <td width="46" height="289">&nbsp;</td>
    <td width="713" valign="top"><table width="635" height="70" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td width="635" height="70" valign="top">
		
		
<%
If(rs_sale.Eof)Then						'判断是否有商品信息
	Response.Write("暂无商品信息!")		'动态输出信息
Else
%>
<table width="656"  border="0" cellspacing="0" cellpadding="0" align="center">
 <tr>
 
 <td width="787" class="tableBorder"><table width="99%"  border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td width="69%" height="376" valign="top"><table width="97%"  border="0" align="center" cellpadding="0" cellspacing="0">
            <tr align="center" valign="top">
<td height="134"><table width="97%" height="355"  border="0" align="center" cellpadding="0" cellspacing="0">
<tr>
<%
		'分页
	  rs_sale.pagesize=6				'设置每页显示记录的数目
	  page1=clng(request("page1"))		'将获取到的记录数转换为整数
	  if page1<1 then page1=1			'为page变量设置默认值
	  rs_sale.absolutepage=page1		'在进行分页显示时，指定指针所在的页
	  for i=1 to rs_sale.pagesize		'循环显示记录信息
	  if i mod 3=0 then					'分栏显示商品信息

%>
<td height="180" valign="top"><table width="25%"  border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="95" align="center"><img src="images/goods/<%=rs_sale("picture")%>" width="150" height="90" border="1"></td>
                    </tr>
                    <tr>
                      <td height="18" align="center"><%=rs_sale("goodsname")%></td>
                    </tr>
                    <tr>
<td height="18" align="center" style="text-decoration:line-through;color:#FF0000">原价：<%=rs_sale("price")%> </td>
                    </tr>
                    <tr>
                      <td height="18" align="center">现价：<%=rs_sale("nowprice")%></td>
                    </tr>
                    <tr>
                      <td height="31" align="center"><a href="#" onClick="javascript:window.open('chakan.asp?id=<%=rs_sale("ID")%>&type=new','','width=415,height=299,toolbar=no,location=no,status=no,menubar=no')"><img src="images/new11.gif" width="72" height="24" border="0"></a>&nbsp;<img src="images/new12.gif" width="57" height="24" onClick="window.location.href='cart_add.asp?goodsID=<%=rs_sale("ID")%>'"></td>
                    </tr>
                  </table></td>
</tr> 
<%else%>
<td height="175" valign="top"><table width="25%"  border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td height="95" align="center"><img src="images/goods/<%=rs_sale("picture")%>" width="150" height="90" border="1"></td>
                      </tr>
                      <tr>
                        <td height="18" align="center"><%=rs_sale("goodsname")%></td>
                      </tr>
                      <tr>
                        <td height="18" align="center" style="text-decoration:line-through;color:#FF0000">原价：<%=rs_sale("price")%> </td>
                      </tr>
                      <tr>
                        <td height="18" align="center">现价：<%=rs_sale("nowprice")%></td>
                      </tr>
                      <tr>
                        <td height="22" align="center"><a href="#" onClick="javascript:window.open('chakan.asp?id=<%=rs_sale("ID")%>&type=new','','width=415,height=299,toolbar=no,location=no,status=no,menubar=no')"><img src="images/new11.gif" width="72" height="24" border="0"></a><a href="#" onClick="javascript:window.open('chakan.asp?id=<%=rs_sale("ID")%>&type=new','','width=430,height=380,toolbar=no,location=no,status=no,menubar=no')"></a> &nbsp;<img src="images/new12.gif" width="57" height="24" onClick="window.location.href='cart_add.asp?goodsID=<%=rs_sale("ID")%>'"></td>
                      </tr>
                    </table></td>
<%
end if
rs_sale.movenext				'向下移动记录指针
if rs_sale.eof then exit for	'退出循环
next
%> 
</table></td>
</tr>
<tr>
<td height="18" align="right"><table width="453" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td height="18"><div align="right" class="STYLE1">
                      <% if page1<>1 then%>
                      <a href=<%=path1%>?page1=1&link=xin>第一页</a> <a href=<%=path1%>?page1=<%=(page1-1)%>&link=xin> 上一页</a>
                      <%end if %>
                      <%if page1<>rs_sale.pagecount then%>
                      <a href=<%=path1%>?page1=<%=(page1+1)%>&link=xin>下一页</a> <a href=<%=path1%>?page1=<%=rs_sale.pagecount%>&link=xin >最后一页</a>
                      <%end if%>
                  </div></td>
                </tr>
              </table></td>
              </tr>
          </table>
          </td>
        </tr>
    </table></td>
  </tr>
</table>
<%end if%>
	</td>
      </tr>
    </table></td>
    <td width="41">&nbsp;</td>
  </tr>
  <tr>
    <td height="13" colspan="3">&nbsp;</td>
  </tr>
  <tr>
    <td height="58" colspan="3">&nbsp;</td>
  </tr>
</table>
</body>
</html>