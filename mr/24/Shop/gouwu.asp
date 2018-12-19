<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
.style1 {color: #f2ab5b}
-->
</style>
<!--#include file="include/conn.asp" -->
<!--#include file="include/include.asp" -->
<!--#include file="chk_user.asp" -->
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
    <td width="590" valign="top"><table border="0" cellpadding="0" cellspacing="0" width="590">
      <tr>
        <td colspan="3"><img name="index_7_r1_c1" src="images/gwc.jpg" width="590" height="49" border="0" alt=""></td>
      </tr>
      <tr>
        <td colspan="3"><img name="index_7_r2_c1" src="images/index_7_r2_c1.jpg" width="590" height="9" border="0" alt=""></td>
      </tr>
      <tr>
        <td width="16" height="643" background="images/index_7_r3_c1.jpg">&nbsp;</td>
        <td width="565" valign="top">
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<%
if request("ProductList")="ProductList" then	'清空购物车
	Session("ProductList")=""
	response.Write("<script>alert('您的购物车为空!');window.location.href='index.asp';</script>")
end if

ProductList = Session("ProductList")	'取得 Session 中的值（N个商品 ID） 赋值给变量 ProductList
Products = Split(Request("Prodid"), ",")	'以逗号分割,赋值给变量 Products （此时变量 Products 以数组形式存在）
For I=0 To UBound(Products)	'按数组的最大下标进行循环
   PutToShopBag Products(I), ProductList	'调用过程并返回参数（商品 ID ，保存商品 ID 的变量 ProductList）
Next
Session("ProductList") = ProductList	'将处理后的变量 ProductList 的值写入到 Session 里

Sub PutToShopBag( Prodid, ProductList )	'定义过程，只有调用时才可以使用
   If Len(ProductList) = 0 Then	'如果变量 ProductList 的值长度为0（等同与值为空）
      ProductList =Prodid	'将变量 ProductList 赋值为商品 ID ，也就是第一次购物的记录
   ElseIf InStr( ProductList, Prodid ) <= 0 Then	'判断变量 ProductList 里面是否有商品 ID 的存在
      ProductList = ProductList&", "&Prodid &""	'多次购物，将多个商品 ID 以逗号分隔组成一个字符串赋值给变量 ProductList
   End If
End Sub

If Request("update") = "update" Then	'隐藏提交，目的：修改商品及数量
   ProductList = ""	'清空购物车
   Products = Split(Request("ProdId"), ", ")	'取得表单提交的商品 ID 并赋值
   For I=0 To UBound(Products)	'按数组的最大下标进行循环
      PutToShopBag Products(I), ProductList	'调用过程并返回参数（商品 ID ，保存商品 ID 的变量 ProductList）
   Next
   Session("ProductList") = ProductList	'将处理后的变量 ProductList 写入到 Session 里，完成了修改商品及数量的目的
End If

If Len(Session("ProductList")) = 0 Then
	response.Write("<script>alert('您的购物车为空!');window.location.href='index.asp';</script>")
	Response.end
end if
%>
<table width="96%" border="0" cellspacing="0" cellpadding="0" align="center">
<%
Set rs=Server.CreateObject("ADODB.RecordSet") 
strsql="select * from shangpin where ID in ("&Session("ProductList")&") order by ID"	'查询保存在 Session 里的变量 ProductList(商品 ID) 
rs.open strsql,conn,1,1
%>
<tr> <td>
<form action="gouwu.asp" method="POST" name="check">
      <table border="0" cellspacing="1" cellpadding="4" align="center" width="100％" bgcolor="BDBDBC">
            <tr bgcolor="#FFFFFF" height="25" align="center"> 
            <td width="40">编 号</td>
            <td width="300">商 品 名 称</td>
            <td width="40">数量</td>
			 <td width="60">市场价</td>
			 <td width="60">会员价</td>
            <td width="60">成交价</td>
			<td width="70">小 计</td>
          </tr>
<%
Sum = 0	'价格总记
Quatity = 1	'商品数量,初始值为1
Do While Not rs.EOF
	Quatity = Request.Form( "Q_" & rs("ID"))	'接收表单提交的商品数量，使用这种接收方法的目的下面表单部分介绍
	If Quatity <= 0 Then	'判断是否第一次购物
	'商品数量为什么会小于零，我们前面不是定义商品数量初始值为1了吗？这就是变量 Quatity 重复赋值的问题
	'虽然定义过商品的数量，但是我们又接收表单提交的商品数量，如果是第一次购买商品的话变量 Quatity 不会在接收表单时被赋予任何值
		Quatity = Session(rs("ID"))	'对应变量 Quatity 进行赋值（以前存储的商品数量）
		If Quatity <= 0 Then Quatity = 1	'如果该商品是用户第一次购买，数量为1
	End If
	Session(rs("ID")) = Quatity	'将商品数量存入 Session 里
	Sum = Sum + rs("huiyuan")*Quatity	'累加器的效果。(新价格总记=旧价格总记+商品价格*商品数量)
%> 
          <tr bgcolor="#FFFFFF" height="25"align="center"> 
            <td><input type="CheckBox" name="ProdId" value="<%=rs("ID")%>" Checked></td>
			 <input type="hidden" name="shuliang" value="<%response.Write Quatity	'商品数量%>">
            <td align="left">&nbsp;<a href="lookpro.asp?ID=<%=rs("ID")%>" target="_blank"><%=rs("mingcheng")%></a></td>
            <td><input type="Text" name="<%response.Write("Q_" & rs("ID")) '下行介绍%>" value="<%response.Write Quatity	'商品数量%>" size="2" class="form"></td>
<%'在我们不知道有多少记录的时候（用户购物的商品，因此无法确定）利用这种方法定义 name 的名称（Q_商品ID）
'有多少商品就会有多少这类的 name 名称，这样 name 不会重复，接收时只要按照这种规律就可以一一对应%>
			<td><%=rs("shichang")%></td>
			<td><%=rs("huiyuan")%></td>
			<td><%=rs("huiyuan")%></td>
			<td><%response.Write(rs("huiyuan")*Quatity)	'价格*数量=该商品的总价格%></td>
          </tr>
		  <input type="hidden" name="xiaoji" value="<%response.Write(rs("huiyuan")*Quatity)	'价格*数量=该商品的总价格%>">
<%
	rs.MoveNext
	Loop
rs.close
set rs=nothing
%> 
<tr bgcolor="#FFFFFF"> 
 <td height="25" colspan="8" align="center" valign="middle">             
                <input type="submit" name="order" value="更新商品"> &nbsp;&nbsp;&nbsp;
                <input type="reset" name="payment" value="去收银台" onClick="window.location.href='shouyin.asp';"> 
                &nbsp;&nbsp;&nbsp; 
				&nbsp;&nbsp;<a href="gouwu.asp?ProductList=ProductList">清空购物车</a>&nbsp;&nbsp;&nbsp;&nbsp;总计：<%=Sum%>
                <input type="hidden" name="update" value="update">
</td>
</tr>
      </table>
      </form>
 </td>
</tr>
      </table>
	  		</td>
        <td width="10" background="images/index_7_r3_c3.jpg">&nbsp;</td>
      </tr>
      <tr>
        <td colspan="3"><img name="index_7_r4_c1" src="images/index_7_r4_c1.jpg" width="590" height="7" border="0" alt=""></td>
      </tr>
    </table></td>
    <td width="8"><img name="index_r2_c3" src="images/index_r2_c3.jpg" width="5" height="753" border="0" alt=""></td>
  </tr>
  <tr>
    <td colspan="3"><!--#include file="foot.asp" --></td>
  </tr>
</table>