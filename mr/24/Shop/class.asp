<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
.style1 {color: #f2ab5b}
-->
</style>
<!--#include file="include/conn.asp" -->
<!--#include file="include/include.asp" -->

<table width="96%" height="153"  border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
<%
if request("id")<>"" then
sql="select * from [shangpin] where classid="&request("id")&" order by id desc"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,1
if rs.eof And rs.bof then
	Response.Write "<p align='center' class='contents'> �÷�����������Ʒ��</p>"
else
	rs.pagesize=8
	SafeRequest(request("page"))
	page=clng(request("page"))
	if page<1 then page=1
	if page>rs.pagecount then page=rs.pagecount
	show rs,page
	sub show(rs,page)
	rs.absolutepage=page
	for i=1 to rs.pagesize
%>
              <td height="89"><table width="255"  border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="130" rowspan="6"><div align="center"><a href="lookpro.asp?id=<%=rs("id")%>" target="_blank"><img src="upfile/<%=rs("tupian")%>" width="110" height="130" border="0"></a></div></td>
                    <td width="20" height="16">&nbsp;</td>
                    <td width="113"><font color="EF9C3E">��<%=rs("mingcheng")%>��</font></td>
                  </tr>
                  <tr>
                    <td height="16">&nbsp;</td>
                    <td><font color="910800">���г��ۣ�<%=rs("shichang")%>��</font></td>
                  </tr>
                  <tr>
                    <td height="16">&nbsp;</td>
                    <td><font color="DD4679">����Ա�ۣ�<%=rs("huiyuan")%>��</font></td>
                  </tr>
                  <tr>
                    <td height="16">&nbsp;</td>
                    <td><a href="lookpro.asp?id=<%=rs("id")%>" target="_blank">���鿴��Ϣ</a>��</td>
                  </tr>
                  <tr>
                    <td height="16">&nbsp;</td>
                    <td>��<a href="#" onClick='javascript:parent.window.location.href="gouwu.asp?ProdId=<%=rs("id")%>&class=class";'>������Ʒ</a>��</td>
                  </tr>
                  <tr>
                    <td height="16">&nbsp;</td>
                    <td><font color="13589B">�����������<%=rs("cishu")%>��</font></td>
                  </tr>
              </table></td>
	<%if i mod 2=0 then%>
            </tr>
<tr><td height="10"></td>
</tr>
<%
end if
	rs.movenext
	if rs.eof then exit for
	next
	end sub
%>
      <tr>
	  <form action='' method='get' name='form'>
        <td height="30" colspan="2">
          <div align="center">
<%	
	if page<>1 then
		response.Write("&nbsp;&nbsp;<a href="&path&"?page=1&id="&request("id")&">��һҳ</a>")
		response.Write("&nbsp;&nbsp;<a href="&path&"?page="&(page-1)&"&id="&request("id")&">��һҳ</a>")
	end if 
	response.Write("&nbsp;&nbsp;��ǰ <font color='#FF0000'>"&page&"</font> ҳ")
	response.Write("&nbsp;&nbsp;�� <font color='#FF0000'>"&rs.recordcount&"</font>/<font color='#FF0000'>"&rs.pagecount&"</font> ҳ")
	if page<>rs.pagecount then
		response.Write("&nbsp;&nbsp;<a href="&path&"?page="&(page+1)&"&id="&request("id")&">��һҳ</a>")
		response.Write("&nbsp;&nbsp;<a href="&path&"?page="&rs.pagecount&"&id="&request("id")&">��ĩҳ</a>")
	end if
	response.Write("&nbsp;&nbsp;��ת��<input type='text' size='2' name='page'>ҳ<input type='hidden' name='id' value='"&request("id")&"'><input type='submit' value='GO'>")
end if
rs.close
set rs=nothing
else
	response.Write("��ѡ������ѯ��")
end if
%>
		    </div></td>
        </form></tr>
        </table>