<!--#include file="conn/conn.asp"--><%'�������ݿ������ļ�%>
<%
Set rs_sale=server.CreateObject("ADODB.RecordSet")                  	  '������¼��
sql="select top 3 * from tb_goods where sale=1 order by INTime desc"	  '��ѯ����
rs_sale.open sql,conn,1,3												  '�򿪼�¼��
%>
<%
If(rs_sale.Eof)Then															'�ж��Ƿ��м�¼��Ϣ
	Response.Write("������Ʒ��Ϣ!")											'������ʾ��Ϣ
Else
%>
<style type="text/css">
<!--
.STYLE2 {
	font-size: 9pt;
	color: #000000;
}
.STYLE4 {font-size: 9pt}
body {
	margin-bottom: 0px;
}
-->
</style>
<br><br><br>
<table width="500" border="0" cellpadding="0" cellspacing="0" align="center">
<tr>
<%
for i=1 to rs_sale.recordcount		'ͳ�Ƽ�¼����
if i mod 3=0 then					'������ʾ��Ʒ��Ϣ
%>
<td height="185" valign="top"><table width="82%"  border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="95" align="center"><img src="images/goods/<%=rs_sale("picture")%>" width="150" height="100" border="1"></td>
                    </tr>
                    <tr>
                      <td height="18" align="center" class="STYLE2"><%=rs_sale("goodsname")%></td>
                    </tr>
                    <tr>
<td height="18" align="center" style="text-decoration:line-through;color:#FF0000" class="STYLE3">ԭ�ۣ�<%=rs_sale("price")%> </td>
                    </tr>
                    <tr>
                      <td height="18" align="center" class="STYLE2">�ּۣ�<%=rs_sale("nowprice")%></td>
                    </tr>
                    <tr>
                      <td height="29" align="center"><a href="#" onClick="javascript:window.open('chakan.asp?id=<%=rs_sale("ID")%>&type=new','','width=415,height=299,toolbar=no,location=no,status=no,menubar=no')"><img src="images/new8.gif" width="72" height="26" border="0" /></a> &nbsp;<img src="images/mai.gif" width="57" height="24" onclick="window.location.href='cart_add.asp?goodsID=<%=rs_sale("ID")%>'" /></td>
                    </tr>
    </table></td>
</tr> 
<%else%>
<td height="185" valign="top"><table width="82%"  border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td height="95" align="center"><img src="images/goods/<%=rs_sale("picture")%>" width="150" height="100" border="1"></td>
                      </tr>
                      <tr>
                        <td height="18" align="center" class="STYLE2"><%=rs_sale("goodsname")%></td>
                      </tr>
                      <tr>
                        <td height="18" align="center" class="STYLE3" style="text-decoration:line-through;color:#FF0000">ԭ�ۣ�<%=rs_sale("price")%> </td>
                      </tr>
                      <tr>
                        <td height="18" align="center" class="STYLE2">�ּۣ�<%=rs_sale("nowprice")%></td>
                      </tr>
                      <tr>
                        <td height="29" align="center"><a href="#" onClick="javascript:window.open('chakan.asp?id=<%=rs_sale("ID")%>&type=new','','width=415,height=299,toolbar=no,location=no,status=no,menubar=no')"><img src="images/new8.gif" width="72" height="26" border="0" /></a> &nbsp;<img src="images/mai.gif" width="57" height="24" onclick="window.location.href='cart_add.asp?goodsID=<%=rs_sale("ID")%>'" /></td>
                      </tr>
    </table></td>
<%
end if
rs_sale.movenext		'�����ƶ���¼ָ��
next
%> 
</table>
<%end if%>