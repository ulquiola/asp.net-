<!--#include file="top.asp"--><!--������վ��ͷ�ļ�-->
<!-- #include file="Conn/conn.asp" --><!--���ݿ������ļ�-->
<%
if request.QueryString("stype")="" then		'��ȡͶƱ������
	if Request.ServerVariables("REMOTE_ADDR")=request.cookies("IPAddress") then
		'��ȡ�û���IP��ַ�����жϸ�IP��ַ�Ƿ��Ѿ�Ͷ��Ʊ
		response.write"<SCRIPT language=JavaScript>alert('��л����֧�֣����Ѿ�Ͷ��Ʊ�ˣ������ظ�ͶƱ��лл��');"
		'������ʾ��Ϣ�Ի���
		response.write"history.back();</SCRIPT>"
		'���ص�ǰһҳ��
	else
		options=request.form("Options")		'��ȡͶƱѡ��
		response.cookies("IPAddress")=Request.ServerVariables("REMOTE_ADDR") 	'��ȡIP��ַ
		conn.execute("update tb_Vote set Option"&options&"=Option"&options&"+1") '����ָ���ļ�¼
		conn.close																 '�رռ�¼��
		set conn=nothing														 '�ͷ�ռ�õ�ϵͳ��Դ�ռ�
	end if
end if
select1="1.�ǳ�ֵ������"															'Ϊ������ֵ 
select2="2.�ȽϿ���"																'Ϊ������ֵ
select3="3.ûʲô�о�"																	
select4="4.��ȫ������"
%>
<style type="text/css">
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-bottom: 0px;
}
.STYLE3 {color: #000073}
</style>
<body>
<table width="800" height="526" border="0" align="center" cellpadding="0" cellspacing="0" background="images/new18.gif">
  <tr>
    <td height="86" colspan="3">&nbsp;</td>
  </tr>
  <tr>
    <td width="46" height="289">&nbsp;</td>
    <td width="713" valign="top"><table width="677" height="70" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td width="677" height="70">
		
		<table width="677" height="65" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="7" height="5" background="images/00.gif">&nbsp;</td>
    <td width="697"><table width="648" border="0" cellspacing="0" cellpadding="0" align="center"  class="wenbenkuang">
<tr> 
<td width="648" height="29" class="table-xiayou"> 
<table border="0" cellpadding="0" cellspacing="0" width="90%" height="48" style="border-collapse: collapse" align="center">
<!-- #include file="Conn/conn.asp" -->     
<%
	set rs=server.CreateObject("adodb.recordset")		'������¼��
	sql="select * from tb_Vote"							'��ѯ����
	rs.open sql,conn,1,3								'�򿪼�¼��
	while not rs.eof 									'�ж��Ƿ��м�¼
%>
 <tr> 
<td height="48" valign="middle" colspan="4" align=center>&nbsp; <span class="STYLE3" style="font-size: 9pt">��վ���ζ�ͶƱ���</span></td>
   			  </tr>
        			<tr>
          				<td width="30%" valign="top" class="STYLE2">���</td>
          				<td colspan="2" valign="top" class="STYLE2">�ٱȷ�</td>
          				<td width="21%" valign="top" class="STYLE2">����</td>
        			</tr>
<%
		an1=int(rs("Option1"))		'�������ֵ���������
		an2=int(rs("Option2"))
		an3=int(rs("Option3"))
		an4=int(rs("Option4"))
		answer=an1+an2+an3+an4		'���ͶƱ��������
		rs.movenext					'�����ƶ���¼ָ��
	wend
%>
        			<tr valign="middle"> 
          				<td height="25" class="STYLE2"><%=select1%>��</td>
          				<td width="34%" height="25"><img src="images/RSCount.gif" width=<%=an1*100/answer%>% height=20><%'��ͼƬ��һ���İٷֱȽ�������%></td>
       				  <td width="15%" height="25"><div align="center" class="STYLE2"><%=round(an1*100/answer,2)%>%</div><%'Ӧ��round�������ذ�ָ��λ�����������������ֵ%></td>
          				<td class="STYLE2"><%=an1%>��</td>
	    			</tr>
		          	<tr valign="middle"> 
          				<td height="25" class="STYLE2"><%=select2%>��</td>
          				<td height="25"><img src=images/RSCount.gif width=<%=int(an2*100/answer)%>% height=20></td>
       				  <td height="25" class="STYLE2"><div align="center"><%=round(an2*100/answer,2)%>%</div><%'Ӧ��round�������ذ�ָ��λ�����������������ֵ%></td>
          				<td class="STYLE2"><%=an2%>��</td>
					</tr>
		          	<tr valign="middle"> 
						<td height="25" class="STYLE2"><%=select3%>��</td>
          				<td height="25"><img src=images/RSCount.gif width=<%=int(an3*100/answer)%>% height=20></td>
       				  <td height="25"><div align="center" class="STYLE2"><%=round(an3*100/answer,2)%>%</div><%'Ӧ��round�������ذ�ָ��λ�����������������ֵ%></td>
          				<td class="STYLE2"><%=an3%>��</td>
		  			</tr>
		          	<tr valign="middle"> 
          				<td height="25" class="STYLE2"><%=select4%>��</td>
          				<td height="25"><img src=images/RSCount.gif width=<%=int(an4*100/answer)%>% height=20></td>
       				  <td height="25" class="STYLE2"> <div align="center"><%=round(an4*100/answer,2)%>%</div></td>
          				<td class="STYLE2"><%=an4%>��</td>
		  			</tr>
        			<tr>
          				<td colspan="4" class="STYLE2"> ���С�<%=answer%>���˲μ�ͶƱ</td>
        			</tr>
		    </table>
      				<p align="center"><span class="STYLE2">��<a href="index.asp">������ҳ</a>��</span> 
<%      
	set rs=nothing     
	conn.close     
	set conn=nothing
%>
</td>
</tr>
</table>
</td>
<td width="30" background="images/000.gif">&nbsp;</td>
</tr>
</table>
		</td>
      </tr>
    </table></td>
    <td width="41">&nbsp;</td>
  </tr>
  <tr>
    <td height="13" colspan="3">&nbsp;</td>
  </tr>
  <tr>
    <td height="138" colspan="3">&nbsp;</td>
  </tr>
</table>