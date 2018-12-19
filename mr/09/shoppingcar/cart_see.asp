<!--#include file="top.asp"--><!--网站的头文件-->
<!--#include file="conn/conn.asp"--><!--数据库连接文件-->
<%
if session("UserName")="" then			'判断用户名是否为空
	response.Write("<script language='javascript'>alert('请先登录!');window.location.href='index.asp';</script>")
	'弹出提示对话框
else
	if not isarray(session("arr")) then		'判断session("arr")变量是否为数组
	response.Redirect("cart_clear.asp")		'跳转到指定的ASP页面
	end if
end if
%>
<html>
<head>
<title>查看购物车</title>
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
<table width="800" height="526" border="0" align="center" cellpadding="0" cellspacing="0" background="images/new20.gif">
  <tr>
    <td height="86" colspan="3">&nbsp;</td>
  </tr>
  <tr>
    <td width="46" height="289">&nbsp;</td>
    <td width="713" valign="top"><table width="673" height="70" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td width="706" height="70">
<table width="596"  border="0" cellspacing="0" cellpadding="0" align="center">
 <tr>
 <td width="597" class="tableBorder">
<table width="100%" height="50"  border="0" align="center" cellpadding="0" cellspacing="0">
    </table>
      <table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="69%" height="189" valign="top"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td height="80">&nbsp;</td>
              </tr>
            <tr valign="top">
              <td height="134" align="center">
<table width="100%" height="126"  border="0" cellpadding="0" cellspacing="0">
			    <tr>
                  <td valign="top">
                    <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td>
<form method="post" action="cart_modify.asp" name="form1">
<table width="92%" height="48" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr align="center" valign="middle">
        <td height="27" class="tableBorder_B1">编号</td>
        <td height="27" class="tableBorder_B1">商品编号</td>
        <td class="tableBorder_B1">商品名称</td>
        <td height="27" class="tableBorder_B1">单价</td>
        <td height="27" class="tableBorder_B1">数量</td>
        <td height="27" class="tableBorder_B1">金额</td>
        <td class="tableBorder_B1">退回</td>
      </tr> 
<%
if isarray(session("arr")) then'判断session("arr")变量是否为数组
arr=session("arr")'为arr变量赋值
	For I = 0 To ubound(arr,1)'获取数组的最大可用下标
		arr_spid=arr(I, 0)'为变量赋值
		arr_dj=arr(I,1)'为变量赋值
		arr_sl=arr(I,2)	'为变量赋值
		if arr_sl<=0 then
			arr_sl=1
		end if
		set arr_rs=Server.CreateObject("ADODB.RecordSet")'创建记录集
		sql="select * from tb_goods where id="&arr_spid		'查询数据
		arr_rs.open sql,conn,1,3							'打开记录集
		if arr_rs.eof and arr_rs.bof then
			response.Write("<script>alert('您的操作有误!');window.location.href='index.asp';</script>")
			session("arr")=""'清空数组
			response.End()
		else
			arr_spname=arr_rs("goodsName")'为arr_spname变量赋值
			sum=sum+arr_dj*arr_sl			'求合计总金额
%>

      <tr align="center" valign="middle"> 
        <td width="32" height="27"><%=i+1%></td>
        <td width="109" height="27"><%=arr_spid%></td> 
        <td width="199" height="27"><%=arr_spname%></td>
        <td width="59" height="27">￥<%=arr_dj%></td> 
        <td width="51" height="27">
          <input name="num<%=i%>" size="7" type="text" class="txt_grey" value="<%=arr_sl%>" onBlur="check(this.form)"></td> 
        <td width="65" height="27">￥<%=(arr_dj*arr_sl)%></td> 
        <td width="34"><a href="cart_move.asp?ID=<%=arr_spid%>"><img src="images/del.gif" width="16" height="16" border="0"></a></td>
        <script language="javascript">
			<!--
			function check(myform)//创建自定义函数
			{
				if(isNaN(myform.num<%=i%>.value) || myform.num<%=i%>.value.indexOf('.',0)!=-1){
					alert("请不要输入非法字符");myform.num<%=i%>.focus();return;}
				if(myform.num<%=i%>.value==""){
					alert("请输入修改的数量");myform.num<%=i%>.focus();return;}	
				myform.submit();
			}
			-->
		</script>
	<%	end if
Next
end if
%>
 </tr>
      </table>
	  </form>
<table width="100%" height="52" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr align="center" valign="middle">
		<td height="10">&nbsp;		</td>
        <td width="24%" height="10" colspan="-3" align="left">&nbsp;</td>
		</tr>
      <tr align="center" valign="middle">
        <td height="21" class="tableBorder_B1">&nbsp;</td>
        <td height="21" colspan="-3" align="left" class="tableBorder_B1">合计总金额：￥<%=sum%></td>
      </tr>
      <tr align="center" valign="middle">
        <td height="21" colspan="2"> <a href="indexs.asp">继续购物</a> | <a href="cart_checkout.asp">去收银台结账</a> | <a href="cart_clear.asp">清空购物车</a> | <a href="#">修改数量</a></td>
        </tr>
</table>						</td>
                      </tr>
                    </table></td>
			     </tr>
              </table>                </td>
            </tr>
            <tr>
              <td height="38" align="right"><a href="sale.asp"><br>
                    <br>
              </a></td>
              </tr>
          </table>          </td>
        </tr>
      </table></td>
  </tr>
</table>
<table width="600"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td></td>
  </tr>
</table>		</td>
      </tr>
    </table></td>
    <td width="41">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="3">&nbsp;</td>
  </tr>
</table>
</body>
</html>