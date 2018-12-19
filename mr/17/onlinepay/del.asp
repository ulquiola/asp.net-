<!--#include file="conn/conn.asp"-->
<%
	orderid = request.QueryString("orderid")
	sql1 = "delete from tb_order_detail where orderID='"&orderid&"'"
	conn.execute(sql1)
	sql2 = "delete from tb_order where orderID = '"&orderid&"'"
	conn.execute(sql2)
%>
<script>
alert('É¾³ý³É¹¦');
location='cart_see.asp';
</script>
<%
	response.End()
%>
