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
if request("ProductList")="ProductList" then	'��չ��ﳵ
	Session("ProductList")=""
	response.Write("<script>alert('���Ĺ��ﳵΪ��!');window.location.href='index.asp';</script>")
end if

ProductList = Session("ProductList")	'ȡ�� Session �е�ֵ��N����Ʒ ID�� ��ֵ������ ProductList
Products = Split(Request("Prodid"), ",")	'�Զ��ŷָ�,��ֵ������ Products ����ʱ���� Products ��������ʽ���ڣ�
For I=0 To UBound(Products)	'�����������±����ѭ��
   PutToShopBag Products(I), ProductList	'���ù��̲����ز�������Ʒ ID ��������Ʒ ID �ı��� ProductList��
Next
Session("ProductList") = ProductList	'�������ı��� ProductList ��ֵд�뵽 Session ��

Sub PutToShopBag( Prodid, ProductList )	'������̣�ֻ�е���ʱ�ſ���ʹ��
   If Len(ProductList) = 0 Then	'������� ProductList ��ֵ����Ϊ0����ͬ��ֵΪ�գ�
      ProductList =Prodid	'������ ProductList ��ֵΪ��Ʒ ID ��Ҳ���ǵ�һ�ι���ļ�¼
   ElseIf InStr( ProductList, Prodid ) <= 0 Then	'�жϱ��� ProductList �����Ƿ�����Ʒ ID �Ĵ���
      ProductList = ProductList&", "&Prodid &""	'��ι���������Ʒ ID �Զ��ŷָ����һ���ַ�����ֵ������ ProductList
   End If
End Sub

If Request("update") = "update" Then	'�����ύ��Ŀ�ģ��޸���Ʒ������
   ProductList = ""	'��չ��ﳵ
   Products = Split(Request("ProdId"), ", ")	'ȡ�ñ��ύ����Ʒ ID ����ֵ
   For I=0 To UBound(Products)	'�����������±����ѭ��
      PutToShopBag Products(I), ProductList	'���ù��̲����ز�������Ʒ ID ��������Ʒ ID �ı��� ProductList��
   Next
   Session("ProductList") = ProductList	'�������ı��� ProductList д�뵽 Session �������޸���Ʒ��������Ŀ��
End If

If Len(Session("ProductList")) = 0 Then
	response.Write("<script>alert('���Ĺ��ﳵΪ��!');window.location.href='index.asp';</script>")
	Response.end
end if
%>
<table width="96%" border="0" cellspacing="0" cellpadding="0" align="center">
<%
Set rs=Server.CreateObject("ADODB.RecordSet") 
strsql="select * from shangpin where ID in ("&Session("ProductList")&") order by ID"	'��ѯ������ Session ��ı��� ProductList(��Ʒ ID) 
rs.open strsql,conn,1,1
%>
<tr> <td>
<form action="gouwu.asp" method="POST" name="check">
      <table border="0" cellspacing="1" cellpadding="4" align="center" width="100��" bgcolor="BDBDBC">
            <tr bgcolor="#FFFFFF" height="25" align="center"> 
            <td width="40">�� ��</td>
            <td width="300">�� Ʒ �� ��</td>
            <td width="40">����</td>
			 <td width="60">�г���</td>
			 <td width="60">��Ա��</td>
            <td width="60">�ɽ���</td>
			<td width="70">С ��</td>
          </tr>
<%
Sum = 0	'�۸��ܼ�
Quatity = 1	'��Ʒ����,��ʼֵΪ1
Do While Not rs.EOF
	Quatity = Request.Form( "Q_" & rs("ID"))	'���ձ��ύ����Ʒ������ʹ�����ֽ��շ�����Ŀ����������ֽ���
	If Quatity <= 0 Then	'�ж��Ƿ��һ�ι���
	'��Ʒ����Ϊʲô��С���㣬����ǰ�治�Ƕ�����Ʒ������ʼֵΪ1��������Ǳ��� Quatity �ظ���ֵ������
	'��Ȼ�������Ʒ�����������������ֽ��ձ��ύ����Ʒ����������ǵ�һ�ι�����Ʒ�Ļ����� Quatity �����ڽ��ձ�ʱ�������κ�ֵ
		Quatity = Session(rs("ID"))	'��Ӧ���� Quatity ���и�ֵ����ǰ�洢����Ʒ������
		If Quatity <= 0 Then Quatity = 1	'�������Ʒ���û���һ�ι�������Ϊ1
	End If
	Session(rs("ID")) = Quatity	'����Ʒ�������� Session ��
	Sum = Sum + rs("huiyuan")*Quatity	'�ۼ�����Ч����(�¼۸��ܼ�=�ɼ۸��ܼ�+��Ʒ�۸�*��Ʒ����)
%> 
          <tr bgcolor="#FFFFFF" height="25"align="center"> 
            <td><input type="CheckBox" name="ProdId" value="<%=rs("ID")%>" Checked></td>
			 <input type="hidden" name="shuliang" value="<%response.Write Quatity	'��Ʒ����%>">
            <td align="left">&nbsp;<a href="lookpro.asp?ID=<%=rs("ID")%>" target="_blank"><%=rs("mingcheng")%></a></td>
            <td><input type="Text" name="<%response.Write("Q_" & rs("ID")) '���н���%>" value="<%response.Write Quatity	'��Ʒ����%>" size="2" class="form"></td>
<%'�����ǲ�֪���ж��ټ�¼��ʱ���û��������Ʒ������޷�ȷ�����������ַ������� name �����ƣ�Q_��ƷID��
'�ж�����Ʒ�ͻ��ж�������� name ���ƣ����� name �����ظ�������ʱֻҪ�������ֹ��ɾͿ���һһ��Ӧ%>
			<td><%=rs("shichang")%></td>
			<td><%=rs("huiyuan")%></td>
			<td><%=rs("huiyuan")%></td>
			<td><%response.Write(rs("huiyuan")*Quatity)	'�۸�*����=����Ʒ���ܼ۸�%></td>
          </tr>
		  <input type="hidden" name="xiaoji" value="<%response.Write(rs("huiyuan")*Quatity)	'�۸�*����=����Ʒ���ܼ۸�%>">
<%
	rs.MoveNext
	Loop
rs.close
set rs=nothing
%> 
<tr bgcolor="#FFFFFF"> 
 <td height="25" colspan="8" align="center" valign="middle">             
                <input type="submit" name="order" value="������Ʒ"> &nbsp;&nbsp;&nbsp;
                <input type="reset" name="payment" value="ȥ����̨" onClick="window.location.href='shouyin.asp';"> 
                &nbsp;&nbsp;&nbsp; 
				&nbsp;&nbsp;<a href="gouwu.asp?ProductList=ProductList">��չ��ﳵ</a>&nbsp;&nbsp;&nbsp;&nbsp;�ܼƣ�<%=Sum%>
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