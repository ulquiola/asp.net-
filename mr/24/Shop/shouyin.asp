<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
.style1 {color: #f2ab5b}
-->
</style>
<!--#include file="include/conn.asp"-->
<!--#include file="chk_user.asp" -->
<!--#include file="include/include.asp" -->
<%
if request("ProductList")="ProductList" then	'��չ��ﳵ
	Session("ProductList")=""
	response.Write("<script>alert('���Ĺ��ﳵΪ��!');window.location.href='index.asp';</script>")
	response.End()
end if
%>
<SCRIPT>
<!--
function chk()
{
   if(document.receiveaddr.shoujianren.value=="") 
	{
		document.receiveaddr.shoujianren.focus();
		alert("����д�ջ���������");
		return false;
	}
	
   if(document.receiveaddr.dizhi.value=="") 
	{
		document.receiveaddr.dizhi.focus();
		alert("����д�ջ��˵�ַ��");
		return false;
	}
	
   if(document.receiveaddr.youbian.value=="") 
	{
		document.receiveaddr.youbian.focus();
		alert("����д�ʱ࣡");
		return false;
	}

   if(document.receiveaddr.tel.value=="") 
	{
		document.receiveaddr.tel.focus();
		alert("����д��ϵ�绰��");
		return false;
	}

   if(document.receiveaddr.mail.value=="") 
	{
		document.receiveaddr.mail.focus();
		alert("����д�����ʼ���");
		return false;
	}

}
//-->
</script> 
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
        <td colspan="3"><img name="index_7_r1_c1" src="images/syt.jpg" width="590" height="49" border="0" alt=""></td>
      </tr>
      <tr>
        <td colspan="3"><img name="index_7_r2_c1" src="images/index_7_r2_c1.jpg" width="590" height="9" border="0" alt=""></td>
      </tr>
      <tr>
        <td width="16" height="643" background="images/index_7_r3_c1.jpg">&nbsp;</td>
        <td width="565" valign="top"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td>&nbsp;</td>
          </tr>
        </table>
<%
if request("action")="add" then
	sql="select * from [dingdan]"
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,3,3
	rs.addnew
	rs("name")=session("user")	'��¼�ö����������û���
	rs("shangpin")=request("ProdId")	'��Ʒ ID����ʱ��Ʒ ID �����ַ�����ʽ����,ÿ����Ʒ ID �Զ��ŷָ���
	rs("shijian")=now()
	rs("shuliang")=request("shuliang")	'��Ʒ��������ʱ��Ʒ���������ַ�����ʽ����,ÿ����Ʒ�����Զ��ŷָ���
	rs("mail")=request("mail")
	rs("dizhi")=trim(request("dizhi"))
	rs("youbian")=trim(request("youbian"))
	rs("zhuangtai")="0"	'������״̬ 0 ���¶��� 1 ���������յ��� 3 �������ѷ��� 4 �û����ջ�
	rs("tel")=trim(request("tel"))
	rs("shoujianren")=trim(request("shoujianren"))
	rs("zhifu")=trim(request("zhifu"))
	rs("songhuo")=trim(request("songhuo"))
	rs("didanhao")=GetOrderNo(Now())	'����Ψһ�Ķ����� GetOrderNo ������ include/include.asp �ļ���
	rs("liuyan")=trim(request("liuyan"))
	rs.update
	rs.close
	set rs=nothing
	response.Write("<script>alert('�����ɹ���\n\n �μ����ĵĶ�����"&GetOrderNo(Now())&"');window.location.href='shouyin.asp?ProductList=ProductList';</script>")
end if
%>
<form name="receiveaddr" method="post" action="shouyin.asp" onSubmit="return chk();">
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
<%
Set rs=Server.CreateObject("ADODB.RecordSet") 
strsql="select * from shangpin where ID in ("&Session("ProductList")&") order by ID"
rs.open strsql,conn,1,1
%>
<tr> <td>
      <table border="0" cellspacing="1" cellpadding="4" align="center" width="96%" bgcolor="BDBDBC">
            <tr bgcolor="#FFFFFF" height="25" align="center"> 
            <td width="45">�� ��</td>
            <td width="194">�� Ʒ �� ��</td>
            <td width="36">����</td>
			 <td width="54">�г���</td>
			 <td width="50">��Ա��</td>
            <td width="49">�ɽ���</td>
			<td width="49">С ��</td>
          </tr>
<%
Sum = 0
Quatity = 1
Do While Not rs.EOF
	i=i+1
	Quatity = Request.Form( "Q_" & rs("ID"))
	If Quatity <= 0 Then 
		Quatity = Session(rs("ID"))
		If Quatity <= 0 Then Quatity = 1
	End If
	Session(rs("ID")) = Quatity
	Dim YourPrice,strprice2,strvipprice,ProScore
	YourPrice= rs("huiyuan")*LngDiscount
	Sum = Sum + (YourPrice + rs("huiyuan"))*Quatity
%> 
          <tr bgcolor="#FFFFFF" height="25"align="center"> 
            <td><input type="CheckBox" name="ProdId" value="<%=rs("ID")%>" Checked></td>
			 
            <td align="left">&nbsp;<a href="product.asp?ID=<%=rs("ID")%>" target="_blank"><%=rs("mingcheng")%></a></td>
            <td><input type="Text" name="<%="Q_" & rs("ID")%>" value="<%=Quatity%>" size="2" class="form"></td>
			<td><%=rs("shichang")%></td>
			<td><%=rs("huiyuan")%></td>
			<td><%=rs("huiyuan")%></td>
			<td><%=(YourPrice + strvipprice)*Quatity%></td>
          </tr>
			 <input type="hidden" name="shuliang" value="<%=Quatity%>">
<%
	rs.MoveNext
	Loop
rs.close
set rs=nothing
%> 
<tr bgcolor="#FFFFFF"> 
 <td height="25" colspan="8" align="center" valign="middle">
<input type="hidden" name="cmdShow" value="Yes">&nbsp;&nbsp;�ܼ۸�<%=Sum%></td>
</tr>
      </table>
 </td>
</tr>
  </table>
<%
sql="select * from [user] where name='"&session("user")&"';"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
if not rs.eof then
	xingming=rs("xingming")
	dizhi=rs("dizhi")
	tel=rs("tel")
	mail=rs("mail")
	youbian=rs("youbian")
end if
%>
<table width="96%" border="0" cellspacing="0" cellpadding="0" align="center" class="table-zuoyou"> 
<tr><td>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="BDBDBC">
          <tr bgcolor=#ffffff> 
		    <input type=hidden name=realname value=timesshop>
            <td width="150">�ջ���������</td>
            <td width="600" height="28">  
              <input name="shoujianren" type="text" size="12" value=<%=xingming%>>
            </td>
          </tr>
          <tr bgcolor=#ffffff> 
            <td>��ϸ��ַ��</td>
            <td height="28"><input name="dizhi" type="text" size="40" value=<%=dizhi%>></td>
          </tr>
          <tr bgcolor=#ffffff> 
            <td>�ʡ����ࣺ</td>
            <td height="28"><input name="youbian" type="text"  size="10" value=<%=youbian%>></td>
          </tr>
          <tr bgcolor=#ffffff> 
            <td>�硡������</td>
            <td height="28"><input name="tel" type="text" size="12" value=<%=tel%>></td>
          </tr>
          <tr bgcolor=#ffffff> 
            <td>�����ʼ���</td>
            <td height="28"><input name="mail" type="text" value=<%=mail%>></td>
          </tr>
          <tr bgcolor=#ffffff> 
            <td>�ͻ���ʽ��</td>
            <td height="28">
          <select name=songhuo size=3 id=deliverymethord>
            <option value=��ͨƽ��>��ͨƽ��</option><option value=�ؿ�ר��>�ؿ�ר�ݣ�EMS��</option><option value=�ͻ����� selected>�ͻ�����</option></select></td>
          </tr>
          <tr bgcolor=#ffffff> 
            <td>֧����ʽ��</td>
            <td height="20">
           <select name=zhifu size=4 id=paymethord><option value=�������л��>�������л��</option><option value=�������л�� selected>�������л��</option><option value=�ʾֻ��>�ʾֻ��</option><option value=��ͨ���л��>��ͨ���л��</option></select>
            </td>
          </tr>
          <tr bgcolor=#ffffff> 
            <td valign="top">�����ԣ�</td>
            <td height="28"><textarea name="liuyan" cols="40" rows="5" id="comments"></textarea></td>
          </tr>
          <tr bgcolor=#ffffff> 
		    <td></td>
            <td><input type="submit" name="Submit3" value="�ύ����"></td>
          </tr>
		 <input type="hidden" name="action" value="add">
      </table> </td>
  </tr> </table></form>
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