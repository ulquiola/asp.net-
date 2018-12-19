<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/conn.asp"--><%'包含数据库连接文件%>
<%
Set rs_sale=server.CreateObject("ADODB.RecordSet")					'创建记录集
sql="select * from tb_goods where newGoods=1 order by INTime desc"	'查询数据
rs_sale.open sql,conn,1,3											'打开记录集
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<style type="text/css">
<!--
.STYLE3 {font-size: 9pt; color: #000000; }
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
<table width="800" height="525" border="0" align="center" cellpadding="0" cellspacing="0" background="images/new13.gif">
  <tr>
    <td height="86" colspan="3">&nbsp;</td>
  </tr>
  <tr>
    <td width="46" height="289">&nbsp;</td>
    <td width="713" valign="top"><table width="635" height="70" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td width="635" height="70" valign="top">
		
		
		<%
If(rs_sale.Eof)Then						'判断是否有商品信息
	Response.Write("暂无商品信息!")		'输出提示信息
Else
%>
        <table width="598"  border="0" cellspacing="0" cellpadding="0" align="center">
 <tr>
 
 <td class="tableBorder"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="69%"  valign="top"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
           <tr align="center" valign="top">
<td  colspan="2"><table width="100%"  border="0" cellpadding="0" cellspacing="0">
<tr>
<%
		'分页
	  rs_sale.pagesize=6'设置每页显示记录的数目
	  page1=clng(request("page1"))'将获取到的记录数转换为整数
	  if page1<1 then page1=1'为page变量设置默认值
		  rs_sale.absolutepage=page1'在进行分页显示时，指定指针所在的页
	  for i=1 to rs_sale.pagesize'循环显示记录信息
	if i mod 3=0 then'分栏显示商品信息

%>
<td  valign="top"><table width="25%"  border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td  align="center"><img src="images/goods/<%=rs_sale("picture")%>" width="150" height="90" border="1"></td>
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
                      <td height="31" align="center"><a href="#" onclick="javascript:window.open('chakan.asp?id=<%=rs_sale("ID")%>&type=new','','width=415,height=299,toolbar=no,location=no,status=no,menubar=no')"><img src="images/new11.gif" width="72" height="24" border="0" /></a>&nbsp;<img src="images/new12.gif" width="57" height="24" onClick="window.location.href='cart_add.asp?goodsID=<%=rs_sale("ID")%>'"></td>
                    </tr>
                  </table></td>
</tr> 
<%else%>
<td height="175" valign="top"><table width="25%"  border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td align="center"><img src="images/goods/<%=rs_sale("picture")%>" width="150" height="90" border="1"></td>
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
                        <td align="center"><a href="#" onClick="javascript:window.open('chakan.asp?id=<%=rs_sale("ID")%>&type=new','','width=415,height=299,toolbar=no,location=no,status=no,menubar=no')"><img src="images/new11.gif" width="72" height="24" border="0"></a><a href="#" onClick="javascript:window.open('chakan.asp?id=<%=rs_sale("ID")%>&type=new','','width=430,height=380,toolbar=no,location=no,status=no,menubar=no')"></a> &nbsp;<img src="images/new12.gif" width="57" height="24" onClick="window.location.href='cart_add.asp?goodsID=<%=rs_sale("ID")%>'"></td>
                      </tr>
                    </table></td>
<%
end if
rs_sale.movenext					'向下移动记录指针
if rs_sale.eof then exit for
next
%> 
</table></td>
</tr>
<tr>
<td height="18" colspan="2" align="right"><table width="512" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td><div align="right" class="STYLE1">
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
    <td height="17" colspan="3">&nbsp;</td>
  </tr>
</table>
</body>
</html>
