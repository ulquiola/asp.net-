<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/conn.asp"--><!--���ݿ������ļ�-->
<%
OrderID=request.QueryString("OrderID")						'��ȡ�������
set rs=server.CreateObject("adodb.recordset")				'������¼��
sql="select * from tb_order where OrderID='"&OrderID&"'"	 '��ѯ����
rs.open sql,conn,1,3									 	'�򿪼�¼��
  	if not rs.eof  then 								 
  		truename=rs("truename")								'Ϊ������ֵ
  		address=rs("address")
  		postcode=rs("postcode")								'Ϊ������ֵ
  		tel=rs("tel")
   		pay=rs("pay")
  		carry=rs("carry")									'Ϊ������ֵ
  		orderdate=rs("orderdate")
  		bz=rs("bz")
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
<style type="text/css">
<!--
.STYLE4 {color: #FF0000}
.STYLE6 {color: #FF0000; font-size: 9pt; }
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
<table width="800" height="528" border="0" align="center" cellpadding="0" cellspacing="0" background="images/new17.gif">
  <tr>
    <td height="86" colspan="3">&nbsp;</td>
  </tr>
  <tr>
    <td width="46" height="289">&nbsp;</td>
    <td width="712" valign="top"><table width="693" height="70" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td width="693" height="70">
		
		<table width="381" height="300" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="376" valign="middle" background="../images/0big.gif"><table width="94%" height="273" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td height="231" valign="top"></div>
          <table width="357" height="230" border="0" align="center" cellpadding="0" cellspacing="0" background="../images/hbig.gif">
          <tr>
            <td width="28%" height="17" valign="bottom"><div align="center" class="STYLE2">��ʵ����</div></td>
            <td width="72%" valign="bottom"><input name="truename" type="text" id="truename" size="25" value="<%=rs("truename")%>"></td>
          </tr>
          <tr>
            <td height="19"><div align="center" class="STYLE2">��ϵ��ַ</div></td>
            <td height="19"><input name="address" type="text" id="address" size="25" value="<%=rs("address")%>" /></td>
          </tr>
          <tr>
            <td height="25"><div align="center" class="STYLE2">��������</div></td>
            <td height="25"><input name="postcode" type="text" id="postcode" size="25" value="<%=rs("postcode")%>" /></td>
          </tr>
          <tr>
            <td height="28"><div align="center" class="STYLE2">��ϵ�绰</div></td>
            <td height="28"><input name="tel" type="text" id="tel" size="25" value="<%=rs("tel")%>" /></td>
          </tr>
          <tr>
            <td height="24"><div align="center" class="STYLE2">���ʽ</div></td>
            <td height="24"><select name="pay" id="pay">
			<option value="���и���"<%if(instr(rs("pay"),"���и���")>0)then%>selected<%end if%>>���и���</option>
			<option value="��������"<%if(instr(rs("pay"),"��������")>0)then%>selected<%end if%>>��������</option>
			<option value="�ֽ�֧��"<%if(instr(rs("pay"),"�ֽ�֧��")>0)then%>selected<%end if%>>�ֽ�֧��</option>
            </select></td>
          </tr>
          <tr>
            <td height="15"><div align="center" class="STYLE2">���ͷ�ʽ</div></td>
            <td height="15"><select name="carry" id="carry">
			<option value="��ͨ�ʼ�"<%if(instr(rs("carry"),"��ͨ�ʼ�")>0)then%>selected<%end if%>>��ͨ�ʼ�</option>
			<option value="�ؿ�ר��"<%if(instr(rs("carry"),"�ؿ�ר��")>0)then%>selected<%end if%>>�ؿ�ר��</option>
			<option value="EMSר�ݷ�ʽ"<%if(instr(rs("carry"),"EMSר�ݷ�ʽ")>0)then%>selected<%end if%>>EMSר�ݷ�ʽ</option>
           </select></td>
          </tr>
          <tr>
            <td height="8"><div align="center" class="STYLE2">��������</div></td>
            <td height="8"><input name="orderdate" type="text" id="orderdate" size="25" value="<%=rs("orderDate")%>" /></td>
          </tr>
          <tr>
            <td height="8"><div align="center" class="STYLE2">�����Ƿ�ִ��</div></td>
            <td height="8">
			<%if rs("enforce")=0 then%>
			<span class="STYLE6">δִ��</span>
			<%end if%>
			<%if rs("enforce")=1 then%>
			<span class="STYLE6">��ִ��</span>
			<%end if%>			</td>
          </tr>
          <tr>
            <td height="26"><div align="center">��&nbsp;&nbsp;&nbsp;ע</div></td>
            <td height="26"><textarea name="bz" cols="30" rows="3" id="bz"><%=rs("bz")%>
</textarea></td>
          </tr>
          <tr>
            <td height="29" colspan="2" valign="bottom"><div align="center"><a href="#" onclick="javascript:window.close()">���رմ��ڡ�</a></div></td>
            </tr>
        </table></td>
      </tr>
      
    </table></td>
  </tr>
</table>
		
		</td>
      </tr>
    </table></td>
    <td width="42">&nbsp;</td>
  </tr>
  <tr>
    <td height="13" colspan="3">&nbsp;</td>
  </tr>
  <tr>
    <td height="129" colspan="3">&nbsp;</td>
  </tr>
</table>
 <% else %>
  <script language="javascript">
  alert("������Ϣ������");				//������ʾ�Ի���
  window.close();						//�رմ���
  </script>
<%end if %>
</body>
</html>
