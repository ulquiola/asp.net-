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
                            <td width="235" bgcolor="#CCCCCC"><div align="center">����ʡ���տƼ����޹�˾-��̴ʵ䶩��</div></td>
                          </tr>
                      </table></td>
                      <td height="25">&nbsp;</td>
                      <td height="25"><%=now()%>&nbsp;</td>
                    </tr>
                    <tr>
                      <td height="1" colspan="4" valign="top"><hr size="1" color="#CCCCCC" width="600"></td>
                    </tr>
                    <tr>
                      <td width="125" height="18"><div align="right">�����ţ�</div></td>
                      <td height="18" colspan="3">&nbsp;<%=rs("orderid")%></td>
                    </tr>
                    <tr>
                      <td height="18"><div align="right">�ջ��ˣ�</div></td>
                      <td width="222" height="18">&nbsp;<%=rs("recuser")%></td>
                      <td width="125" height="18"><div align="right">�ʱࣺ</div></td>
                      <td width="208" height="18">&nbsp;<%=rs("yb")%></td>
                    </tr>
                    <tr>
                      <td height="18"><div align="right">�ƶ��绰��</div></td>
                      <td height="18">&nbsp;<%=rs("mtel")%></td>
                      <td height="18"><div align="right">�̶��绰��</div></td>
                      <td height="18">&nbsp;<%=rs("gtel")%></td>
                    </tr>
                    <tr>
                      <td height="18"><div align="right">��ϵ��ַ��</div></td>
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
                                                              <td height="27" class="tableBorder_B1">���</td>
                                                              <td height="27" class="tableBorder_B1">��Ʒ���</td>
                                                              <td class="tableBorder_B1">��Ʒ����</td>
                                                              <td height="27" class="tableBorder_B1">����</td>
                                                              <td height="27" class="tableBorder_B1">����</td>
                                                              <td height="27" class="tableBorder_B1">���</td>
                                                             
                                                            </tr>
															
<%
if isarray(session("arr")) then'�ж�session("arr")�����Ƿ�Ϊ����
arr=session("arr")'Ϊarr������ֵ
	For I = 0 To ubound(arr,1)'��ȡ������������±�
		arr_spid=arr(I, 0)'Ϊ������ֵ
		arr_dj=arr(I,1)'Ϊ������ֵ
		arr_sl=arr(I,2)	'Ϊ������ֵ
		if arr_sl<=0 then
			arr_sl=1
		end if
		set arr_rs=Server.CreateObject("ADODB.RecordSet")'������¼��
		sql="select * from tb_goods where id="&arr_spid		'��ѯ����
		arr_rs.open sql,conn,1,3							'�򿪼�¼��
		if arr_rs.eof and arr_rs.bof then
			response.Write("<script>alert('���Ĳ�������!');window.location.href='index.asp';</script>")
			session("arr")=""'�������
			response.End()
		else
			arr_spname=arr_rs("goodsName")'Ϊarr_spname������ֵ
			sum=sum+arr_dj*arr_sl			'��ϼ��ܽ��
%>
                                                            <tr align="center" valign="middle">
                                                              <td width="32" height="27"><%=i+1%></td>
                                                              <td width="109" height="27"><%=arr_spid%></td>
                                                              <td width="199" height="27"><%=arr_spname%></td>
                                                              <td width="59" height="27">��<%=arr_dj%></td>
                                                              <td width="51" height="27"><%=arr_sl%></td>
                                                              <td width="65" height="27">��<%=(arr_dj*arr_sl)%></td>
<%	end if
Next
end if
%>
                                                            </tr>
															<tr>
																<td colspan="6" height="27">�ϼ��ܽ���<%=sum%></td>
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