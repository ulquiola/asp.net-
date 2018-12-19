<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="top.asp"-->
<!--#include file="conn/conn.asp"-->
<!--#include file="alipayto/alipay_payto.asp"-->
<%
orderid = request.QueryString("orderid")
sum = request.QueryString("sum")
orderdate = request.QueryString("orderdate")
set rs=server.CreateObject("adodb.recordset")
sql="select * from tb_Order where orderID = '"&orderid&"'"
rs.open sql,conn,1,3
'支付宝
shijian=now()
dingdan=year(shijian)&month(shijian)&day(shijian)&hour(shijian)&minute(shijian)&second(shijian)
   
   subject			=	orderid		'商品名称
	body			=	"测试"		'body			商品描述
	out_trade_no    =   dingdan       
	price		    =	sum				'price商品单价			0.01～50000.00
    quantity        =   "1"               '商品数量,如果走购物车默认为1
	discount        =   "0"               '商品折扣
    seller_email    =    "mingrisoft"   '卖家的支付宝帐号
	Set AlipayObj	= New creatAlipayItemURL
	itemUrl=AlipayObj.creatAlipayItemURL(subject,body,out_trade_no,price,quantity,seller_email)
%>

<table width="800" height="523" border="0" align="center" cellpadding="0" cellspacing="0" background="images/new21.gif">
  <tr>
    <td height="38" colspan="3">&nbsp;</td>
  </tr>
  <tr>
    <td width="46" height="289">&nbsp;</td>
    <td width="713" valign="top"><p>&nbsp;</p>
      <table width="750" height="60" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td><br>
            <table width="630" height="50" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#999999">
              <tr>
                <td bgcolor="#FFFFFF"><table width="630" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="130" height="22"><div align="right">订单号：</div></td>
                      <td width="211">&nbsp;<font color=red><strong><%=orderid%></strong></font></td>
                      <td width="130"><div align="right">需支付金额：</div></td>
                      <td width="159">&nbsp; <font color=red><strong>"<%=sum%>"&nbsp;元</strong></font> </td>
                    </tr>
                    <tr>
                      <td height="30" colspan="4"><table width="500" height="1" border="0" align="center" cellpadding="0" cellspacing="0">
                          <tr>
                            <td bgcolor="#CCCCCC"></td>
                          </tr>
                        </table>
                          <table width="500" height="50" border="0" align="center" cellpadding="0" cellspacing="0">
                            <tr>
                              <td style="line-height:2">&nbsp;<font color="#FF0000">*</font>&nbsp;<font color="#999999">只有在网上支付成功后，我公司才会为您邮递。</font><br>
                                &nbsp;<font color="#FF0000">*</font>&nbsp;<font color="#999999">我们会在48小时内保留您的订单，请及时支付。</font></td>
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
                      <td width="338">&nbsp;</td>
                      <td width="68"><img src="images/sy_17.jpg" width="60" height="25" style="cursor:hand" onClick="javascript:if(window.confirm('如果取消该订单，则该订单将被删除，您需要重新购买！')==true){window.location.href='del.asp?orderid=<%=orderid%>';}"/></td>
                      <!--支付宝支付的接口操作，提交的数据-->
                      <td width="100"><a href="<%=itemUrl%>"><img src="images/sy_19.gif" width="90" height="25" border="0"></a></td>
                      <!--———————工行生成操作函数———————————-->
<%
	Dim bb,rc
	Set bb =CreateObject("ICBCEBANKUTIL.B2CUtil")
	rc=bb.init ("d:\user.crt","d:\user.crt","d:\user.key","11111111")
	if rc=0 then 
		response.write "初始化成功.<br>"
	end if
 	src = "ICBC_PERBANK_B2C1.0.0.00200EC200000120200029109000030106http://www.geticbcmsg.com.cn/servletHS0000000011000010200508011925560"
	
	ssrc = bb.signC(src, Len(src))	

	rc=bb.verifySignC(src, Len(src), ssrc, Len(ssrc)) 

	cert=bb.getCert(1)

%>
                      <!--工行支付的接口操作，提交的数据-->

<form  action="https://mybank.icbc.com.cn/servlet/ICBCINBSEBusinessServlet" method="post" name="form_bank">
    <input name="interfaceName" type="hidden" value="ICBC_PERBANK_B2C"/>
    <input name="interfaceVersion" type="hidden" value="1.0.0.0"/>
       <input name="orderid" type="hidden" value="<%=request.QueryString("orderid")%>"/><!-- 订单号  -->
       <input name="amount" type="hidden" value="<%=request.QueryString("sum")%>"/><!-- 商品总金额，金额以分为单位  -->
    <input name="curType" type="hidden" value="001"/>							<!-- 币种目前只支持人民币，代码为“001” -->
       <input name="merID" type="hidden" value="<%=request.QueryString("merID")%>"/>	<!-- 银行提供 -->
       <input name="merAcct" type="hidden" value="<%=request.QueryString("merAcct")%>"/><!-- 银行提供 -->
    <input name="verifyJoinFlag" type="hidden" value="0"/>							
	<!-- “1”判断该客户是否与商户联名；取值“0”不检验客户是否与商户联名 -->
    <input name="notifyType" type="hidden" value="HS"/>							
	<!-- HS方式实时发送通知；AG方式不发送通知； -->
      <input name="merURL" type="hidden" value="<%=src%>"/>							
	  <!-- 接收银行通知地址，目前只支持http协议80端口 -->
    <input name="resultType" type="hidden"  value="0"/>								
	<!-- 对于HS方式“0”：发送成功或者失败信息；“1”，只发送交易成功信息 -->
       <input name="orderDate" type="hidden" value="<%=request.QueryString("orderDate")%>"/><!-- 订单日期 -->
    <input name="merSignMsg" type="hidden" value="<%=ssrc%>" /><!-- 签名 -->
    <input name="merCert" type="hidden" value="<%=cert%>" /><!-- 证书 -->
</form>
<td width="124"><img src="images/sy_18.jpg" width="70" height="25" onClick="javascript:form_bank.submit();" style="cursor:hand"/></td>
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
