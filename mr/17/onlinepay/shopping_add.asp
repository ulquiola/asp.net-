<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="conn/conn.asp"--><!--数据库连接文件-->
<!--#include file="top.asp"-->
<%
orderID = session("orderID")
set rs=server.CreateObject("adodb.recordset")
sql="select * from tb_Order where orderID = '"&orderID&"'"
rs.open sql,conn,1,3
%>
<table width="800" height="523" border="0" align="center" cellpadding="0" cellspacing="0" background="images/new21.gif">
  <tr>
    <td height="38" colspan="3">&nbsp;</td>
  </tr>
  <tr>
    <td width="46" height="289">&nbsp;</td>
    <td width="713" valign="top"><table width="750" height="60" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td><br>
            <table width="630" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#999999">
              <script language="javascript">
	  function openprintwindow(x,y,z){
	    
	     window.open("print.asp?orderID="+x+"&act="+z,"newframe","top=200,left=200,width=635,height="+(230+27*y)+",menubar=no,location=no,toolbar=no,scrollbars=no,status=no");
	
	  }
	
	  </script>
              <tr>
                <td bgcolor="#FFFFFF"><table width="630" border="0" align="center" cellpadding="0" cellspacing="0">
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
                                        <tr>
                                          <td height="80">&nbsp;</td>
                                        </tr>
                                        <tr valign="top">
                                          <td height="134" align="center"><table width="100%" height="126"  border="0" cellpadding="0" cellspacing="0">
                                              <tr>
                                                <td valign="top"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                                                    <tr>
                                                      <td><form method="post" action="cart_modify.asp" name="form1">
                                                          <table width="92%" height="48" border="0" align="center" cellpadding="0" cellspacing="0">
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
dim gsum
gsum = 0
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
			gsum = gsum + 1
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
                                                          </table>
                                                      </form>
                                                          <table width="100%" height="47" border="0" align="center" cellpadding="0" cellspacing="0">
                                                            <tr align="center" valign="middle">
                                                              <td height="10">&nbsp;</td>
                                                              <td width="24%" height="10" colspan="-3" align="left">&nbsp;</td>
                                                            </tr>
                                                            <tr align="center" valign="middle">
                                                              <td height="21" class="tableBorder_B1">&nbsp;</td>
                                                              <td height="21" colspan="-3" align="left" class="tableBorder_B1">合计总金额：￥<%=sum%></td>
                                                            </tr>
                                                            <tr align="center" valign="middle">
                                                              <td height="13" colspan="2">&nbsp;</td>
                                                            </tr>
                                                        </table></td>
                                                    </tr>
                                                </table></td>
                                              </tr>
                                          </table></td>
                                        </tr>
                                        <tr>
                                          <td height="38" align="right"><a href="sale.asp"><br>
                                                <br>
                                          </a></td>
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
          <br>
            <table width="630" height="25" border="0" align="center" cellpadding="0" cellspacing="1">
              <tr>
                <td><table width="630" height="25" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="340"></td>
                      <td width="75"><img src="images/sy_12.jpg" width="60" height="25" onClick="javascript:openprintwindow('<%=rs("orderid")%>','<%=gsum%>','p')" style="cursor:hand"/></td>
                      <td width="90"><img src="images/sy_14.jpg" width="69" height="25"  onclick="javascript:openprintwindow('<%=rs("orderid")%>','<%=gsum%>','v')" style="cursor:hand"/></td>
                      <td width="125"><img src="images/sy_16.jpg" width="70" height="25" onClick="javascript:window.location.href='shoppingadd_deal.asp?orderid=<%=rs("orderid")%>&sum=<%=sum%>&orderdate=<%=rs("orderdate")%>'" style="cursor:hand"/></td>
                    </tr>
                </table></td>
              </tr>
            </table>
          <br>
        </td>
      </tr>
    </table></td>
    <td width="41">&nbsp;</td>
  </tr>
  <tr>
    <td height="13" colspan="3">&nbsp;</td>
  </tr>
  <tr>
    <td height="77" colspan="3">&nbsp;</td>
  </tr>
</table>
<p>&nbsp; </p>