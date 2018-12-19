<!--#include file="conn/conn.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title></title>
</head>
<script>     
  function printview()     
  {     
  document.all.WebBrowser1.ExecWB(7,1) ;
  window.close();  
  }     
  function prints()     
  {     
 document.all.WebBrowser1.Execwb(6,1)
  window.close();  
  }     
 </script>     
<%
	orderid = request.QueryString("orderid")
	act = request.QueryString("act")
	set rs=server.CreateObject("adodb.recordset")
	sql="select * from tb_Order where orderID = '"&orderid&"'"
	rs.open sql,conn,1,3
%>
<object   ID='WebBrowser1'   WIDTH=0   HEIGHT=0   CLASSID='CLSID:8856F961-340A-11D0-A96B-00C04FD705A2'></object>
 

<body topmargin="0" leftmargin="0" bottommargin="0" onLoad="<% if act="p" then %>prints();<% elseif act="v" then %> printview();<% end if %> ">
<table width="630" border="1" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="25" colspan="2"><table width="250" height="20" border="0" align="left" cellpadding="0" cellspacing="0">
                          <tr>
                            <td width="15" bgcolor="#FFFFFF"><div align="center"></div></td>
                            <td width="235" bgcolor="#CCCCCC"><div align="center">吉林省明日科技有限公司-编程词典订单</div></td>
                          </tr>
                      </table></td>
                      <td height="25">&nbsp;</td>
                      <td height="25"><%=now()%>&nbsp;</td>
                    </tr>
                    <tr>
                      <td height="1" colspan="4" valign="top"><hr size="1" color="#CCCCCC" width="600"></td>
                    </tr>
                    <tr>
                      <td width="125" height="18"><div align="right">订单号：</div></td>
                      <td height="18" colspan="3">&nbsp;<%=rs("orderid")%></td>
                    </tr>
                    <tr>
                      <td height="18"><div align="right">收货人：</div></td>
                      <td width="222" height="18">&nbsp;<%=rs("recuser")%></td>
                      <td width="125" height="18"><div align="right">邮编：</div></td>
                      <td width="208" height="18">&nbsp;<%=rs("yb")%></td>
                    </tr>
                    <tr>
                      <td height="18"><div align="right">移动电话：</div></td>
                      <td height="18">&nbsp;<%=rs("mtel")%></td>
                      <td height="18"><div align="right">固定电话：</div></td>
                      <td height="18">&nbsp;<%=rs("gtel")%></td>
                    </tr>
                    <tr>
                      <td height="18"><div align="right">联系地址：</div></td>
                      <td height="18" colspan="3">&nbsp;<%=rs("address")%></td>
                    </tr>
                    <tr>
                      <td height="20" colspan="4"><br>
                        <table width="596"  border="0" cellspacing="0" cellpadding="0" align="center">
                          <tr>
                            <td width="597" class="tableBorder"><table width="100%" height="50"  border="0" align="center" cellpadding="0" cellspacing="0">
                              </table>
                                <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                                  <tr>
                                    <td width="69%" height="189" valign="top"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                                        <tr valign="top">
                                          <td height="134" align="center"><table width="100%" height="126"  border="0" cellpadding="0" cellspacing="0">
                                              <tr>
                                                <td valign="top"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                                                    <tr>
                                                      <td><form method="post" action="cart_modify.asp" name="form1">
                                                          <table width="92%" height="48" border="1" align="center" cellpadding="0" cellspacing="0">
                                                            <tr align="center" valign="middle">
                                                              <td height="27" class="tableBorder_B1">编号</td>
                                                              <td height="27" class="tableBorder_B1">商品编号</td>
                                                              <td class="tableBorder_B1">商品名称</td>
                                                              <td height="27" class="tableBorder_B1">单价</td>
                                                              <td height="27" class="tableBorder_B1">数量</td>
                                                              <td height="27" class="tableBorder_B1">金额</td>
                                                             
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
                                                              <td width="51" height="27"><%=arr_sl%></td>
                                                              <td width="65" height="27">￥<%=(arr_dj*arr_sl)%></td>
<%	end if
Next
end if
%>
                                                            </tr>
															<tr>
																<td colspan="6" height="27">合计总金额：￥<%=sum%></td>
															</tr>
                                                          </table>
                                                      </form>
                                                          </td>
                                                    </tr>
                                                </table></td>
                                              </tr>
                                          </table></td>
                                        </tr>
                                    </table></td>
                                  </tr>
                              </table></td>
                          </tr>
                        </table></td>
                    </tr>
                </table>
</body>
</html>