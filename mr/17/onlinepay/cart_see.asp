<!--#include file="top.asp"--><!--��վ��ͷ�ļ�-->
<!--#include file="conn/conn.asp"--><!--���ݿ������ļ�-->
<%
if session("UserName")="" then			'�ж��û����Ƿ�Ϊ��
	response.Write("<script language='javascript'>alert('���ȵ�¼!');window.location.href='index.asp';</script>")
	'������ʾ�Ի���
else
	if not isarray(session("arr")) then		'�ж�session("arr")�����Ƿ�Ϊ����
	response.Redirect("cart_clear.asp")		'��ת��ָ����ASPҳ��
	end if
end if
%>
<html>
<head>
<title>�鿴���ﳵ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<style type="text/css">
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-bottom: 0px;
}
</style>
<body>
<table width="800" height="526" border="0" align="center" cellpadding="0" cellspacing="0" background="images/new20.gif">
  <tr>
    <td height="86" colspan="3">&nbsp;</td>
  </tr>
  <tr>
    <td width="46" height="289">&nbsp;</td>
    <td width="713" valign="top"><table width="673" height="70" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td width="706" height="70">
<table width="596"  border="0" cellspacing="0" cellpadding="0" align="center">
 <tr>
 <td width="597" class="tableBorder">
<table width="100%" height="50"  border="0" align="center" cellpadding="0" cellspacing="0">
    </table>
      <table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="69%" height="189" valign="top"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td height="80">&nbsp;</td>
              </tr>
            <tr valign="top">
              <td height="134" align="center">
<table width="100%" height="126"  border="0" cellpadding="0" cellspacing="0">
			    <tr>
                  <td valign="top">
                    <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td>
<form method="post" action="cart_modify.asp" name="form1">
<table width="92%" height="48" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr align="center" valign="middle">
        <td height="27" class="tableBorder_B1">���</td>
        <td height="27" class="tableBorder_B1">��Ʒ���</td>
        <td class="tableBorder_B1">��Ʒ����</td>
        <td height="27" class="tableBorder_B1">����</td>
        <td height="27" class="tableBorder_B1">����</td>
        <td height="27" class="tableBorder_B1">���</td>
        <td class="tableBorder_B1">�˻�</td>
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
        <td width="51" height="27">
          <input name="num<%=i%>" size="7" type="text" class="txt_grey" value="<%=arr_sl%>" onBlur="check(this.form)"></td> 
        <td width="65" height="27">��<%=(arr_dj*arr_sl)%></td> 
        <td width="34"><a href="cart_move.asp?ID=<%=arr_spid%>"><img src="images/del.gif" width="16" height="16" border="0"></a></td>
        <script language="javascript">
			<!--
			function check(myform)//�����Զ��庯��
			{
				if(isNaN(myform.num<%=i%>.value) || myform.num<%=i%>.value.indexOf('.',0)!=-1){
					alert("�벻Ҫ����Ƿ��ַ�");myform.num<%=i%>.focus();return;}
				if(myform.num<%=i%>.value==""){
					alert("�������޸ĵ�����");myform.num<%=i%>.focus();return;}	
				myform.submit();
			}
			-->
		</script>
	<%	end if
Next
end if
%>
 </tr>
      </table>
	  </form>
<table width="100%" height="52" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr align="center" valign="middle">
		<td height="10">&nbsp;		</td>
        <td width="24%" height="10" colspan="-3" align="left">&nbsp;</td>
		</tr>
      <tr align="center" valign="middle">
        <td height="21" class="tableBorder_B1">&nbsp;</td>
        <td height="21" colspan="-3" align="left" class="tableBorder_B1">�ϼ��ܽ���<%=sum%></td>
      </tr>
      <tr align="center" valign="middle">
        <td height="21" colspan="2"> <a href="indexs.asp">��������</a> | <a href="cart_checkout.asp">ȥ����̨����</a> | <a href="cart_clear.asp">��չ��ﳵ</a> | <a href="#">�޸�����</a></td>
        </tr>
</table>						</td>
                      </tr>
                    </table></td>
			     </tr>
              </table>                </td>
            </tr>
            <tr>
              <td height="38" align="right"><a href="sale.asp"><br>
                    <br>
              </a></td>
              </tr>
          </table>          </td>
        </tr>
      </table></td>
  </tr>
</table>
<table width="600"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td></td>
  </tr>
</table>		</td>
      </tr>
    </table></td>
    <td width="41">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="3">&nbsp;</td>
  </tr>
</table>
</body>
</html>