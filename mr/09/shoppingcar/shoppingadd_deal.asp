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
'֧����
shijian=now()
dingdan=year(shijian)&month(shijian)&day(shijian)&hour(shijian)&minute(shijian)&second(shijian)
   
   subject			=	orderid		'��Ʒ����
	body			=	"����"		'body			��Ʒ����
	out_trade_no    =   dingdan       
	price		    =	sum				'price��Ʒ����			0.01��50000.00
    quantity        =   "1"               '��Ʒ����,����߹��ﳵĬ��Ϊ1
	discount        =   "0"               '��Ʒ�ۿ�
    seller_email    =    "mingrisoft"   '���ҵ�֧�����ʺ�
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
                      <td width="130" height="22"><div align="right">�����ţ�</div></td>
                      <td width="211">&nbsp;<font color=red><strong><%=orderid%></strong></font></td>
                      <td width="130"><div align="right">��֧����</div></td>
                      <td width="159">&nbsp; <font color=red><strong>"<%=sum%>"&nbsp;Ԫ</strong></font> </td>
                    </tr>
                    <tr>
                      <td height="30" colspan="4"><table width="500" height="1" border="0" align="center" cellpadding="0" cellspacing="0">
                          <tr>
                            <td bgcolor="#CCCCCC"></td>
                          </tr>
                        </table>
                          <table width="500" height="50" border="0" align="center" cellpadding="0" cellspacing="0">
                            <tr>
                              <td style="line-height:2">&nbsp;<font color="#FF0000">*</font>&nbsp;<font color="#999999">ֻ��������֧���ɹ����ҹ�˾�Ż�Ϊ���ʵݡ�</font><br>
                                &nbsp;<font color="#FF0000">*</font>&nbsp;<font color="#999999">���ǻ���48Сʱ�ڱ������Ķ������뼰ʱ֧����</font></td>
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
                      <td width="68"><img src="images/sy_17.jpg" width="60" height="25" style="cursor:hand" onClick="javascript:if(window.confirm('���ȡ���ö�������ö�������ɾ��������Ҫ���¹���')==true){window.location.href='del.asp?orderid=<%=orderid%>';}"/></td>
                      <!--֧����֧���Ľӿڲ������ύ������-->
                      <td width="100"><a href="<%=itemUrl%>"><img src="images/sy_19.gif" width="90" height="25" border="0"></a></td>
                      <!--���������������������ɲ�����������������������������-->
<%
	Dim bb,rc
	Set bb =CreateObject("ICBCEBANKUTIL.B2CUtil")
	rc=bb.init ("d:\user.crt","d:\user.crt","d:\user.key","11111111")
	if rc=0 then 
		response.write "��ʼ���ɹ�.<br>"
	end if
 	src = "ICBC_PERBANK_B2C1.0.0.00200EC200000120200029109000030106http://www.geticbcmsg.com.cn/servletHS0000000011000010200508011925560"
	
	ssrc = bb.signC(src, Len(src))	

	rc=bb.verifySignC(src, Len(src), ssrc, Len(ssrc)) 

	cert=bb.getCert(1)

%>
                      <!--����֧���Ľӿڲ������ύ������-->

<form  action="https://mybank.icbc.com.cn/servlet/ICBCINBSEBusinessServlet" method="post" name="form_bank">
    <input name="interfaceName" type="hidden" value="ICBC_PERBANK_B2C"/>
    <input name="interfaceVersion" type="hidden" value="1.0.0.0"/>
       <input name="orderid" type="hidden" value="<%=request.QueryString("orderid")%>"/><!-- ������  -->
       <input name="amount" type="hidden" value="<%=request.QueryString("sum")%>"/><!-- ��Ʒ�ܽ�����Է�Ϊ��λ  -->
    <input name="curType" type="hidden" value="001"/>							<!-- ����Ŀǰֻ֧������ң�����Ϊ��001�� -->
       <input name="merID" type="hidden" value="<%=request.QueryString("merID")%>"/>	<!-- �����ṩ -->
       <input name="merAcct" type="hidden" value="<%=request.QueryString("merAcct")%>"/><!-- �����ṩ -->
    <input name="verifyJoinFlag" type="hidden" value="0"/>							
	<!-- ��1���жϸÿͻ��Ƿ����̻�������ȡֵ��0��������ͻ��Ƿ����̻����� -->
    <input name="notifyType" type="hidden" value="HS"/>							
	<!-- HS��ʽʵʱ����֪ͨ��AG��ʽ������֪ͨ�� -->
      <input name="merURL" type="hidden" value="<%=src%>"/>							
	  <!-- ��������֪ͨ��ַ��Ŀǰֻ֧��httpЭ��80�˿� -->
    <input name="resultType" type="hidden"  value="0"/>								
	<!-- ����HS��ʽ��0�������ͳɹ�����ʧ����Ϣ����1����ֻ���ͽ��׳ɹ���Ϣ -->
       <input name="orderDate" type="hidden" value="<%=request.QueryString("orderDate")%>"/><!-- �������� -->
    <input name="merSignMsg" type="hidden" value="<%=ssrc%>" /><!-- ǩ�� -->
    <input name="merCert" type="hidden" value="<%=cert%>" /><!-- ֤�� -->
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
