<!--#include file="conn/conn.asp"--><!--���ݿ������ļ�-->
<!--#include file="top.asp"--><!--������վ��ͷ�ļ�-->
<%
dim cgid
cgid = ""
if request.Form("recuser")<>"" then		'�ж��û����Ƿ�Ϊ��
	if isarray(session("arr")) then			'�ж�session("arr")�����Ƿ�Ϊ����
		arr=session("arr")					'Ϊarr������ֵ
		bnumber=ubound(arr,1)+1				'Ϊbumber������ֵ
		recuser=request.Form("recuser")	'Ϊbumber������ֵ
		address=request.Form("address")		'Ϊbumber������ֵ
		yb=request.Form("yb")	'Ϊbumber������ֵ
		qq=request.Form("qq")
		email=request.Form("email")
		mtel=request.Form("mtel")
		gtel=request.Form("gtel")
		shfs=request.Form("shfs")
		set rs=server.CreateObject("adodb.recordset")
		'����ָ����¼ͨ��with(updlock)��ɡ�
		sql="select orderID from tb_Order with(updlock)"
		rs.open sql,conn,1,3
		'������ת��ΪYYYYMMDD��ʽ���ַ���
		if len(month(date()))=1 then
			cmonth="0"& cstr(month(date()))
		else
			cmonth=cstr(month(date()))
		end if
		if len(day(date()))=1 then
		cday="0"& cstr(day(date()))
		else
		cday=cstr(day(date()))
		end if
		if not rs.eof then 
			set rs_M=server.CreateObject("adodb.recordset")
			sql1="select max(orderID) as orderID from tb_Order"
			rs_M.open sql1,conn,1,3
			str=rs_M("orderID")
			cno = right(cstr(int(mid(str,11,5))+100001),5)
			intno="CG"&cstr(year(date()))&cmonth&cday&cno '����µ����ݱ��
			cgid=intno
		else
			cgid = "CG"&cstr(year(date()))&cmonth&cday&"00000"
		end if
		
		on error resume next				'�������������
		conn.BeginTrans    					'��ʼ����
		sql="insert into tb_Order(OrderID,bnumber,recuser,address,yb,qq,email,mtel,gtel,shfs) values('"&cgid&"',"&bnumber&",'"&recuser&"','"&address&"','"&yb&"','"&qq&"','"&email&"','"&mtel&"','"&gtel&"','"&shfs&"')"		'����û���Ϣ
		conn.execute(sql)																		'ִ��SQL���
		For I = 0 To ubound(arr,1)				
			arr_spid=arr(I, 0)																	'Ϊ���������ֵ
			arr_dj=arr(I,1)																		'Ϊ���������ֵ
			arr_sl=arr(I,2)	
			if arr_sl<=0 then
				arr_sl=1
			end if
			str="insert into tb_order_Detail (orderID,goodsID,price,number) values('"&cgid&"',"&arr_spid&","&arr_dj&","&arr_sl&")"
			conn.execute(str)																	'ִ��SQL���
		next
		if err<>0 then
			conn.RollbackTrans  																'����ع�
			response.Write(err.description)
			%>
			<script language="javascript">
			alert("����ʧ��!");																	//����������ʾ�Ի���
			window.location.href="indexs.asp";													//��ת��ָ����ASPҳ��
			</script>
			<%
			response.End()
			else
				session("orderID") = cgid
			%>
			<script language="javascript">
			alert("������ӳɹ���");							
			window.location.href="shopping_add.asp";
			</script>
			<%
			conn.CommitTrans   
			end if
			%>
		<%
		'�ύ�궩������չ��ﳵ%>
	<%else
		response.Write("<script language='javascript'>alert('���Ĳ�������!');window.location.href='index.asp';</script>")
	end if
else
%>
<html>
<script src="js/chkreg.js"></script>
<head>
<title>����̨����</title>
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
<table width="800" height="513" border="0" align="center" cellpadding="0" cellspacing="0" background="images/new21.gif">
  <tr>
    <td height="38" colspan="3">&nbsp;</td>
  </tr>
  <tr>
    <td width="46" height="289">&nbsp;</td>
    <td width="713" valign="top"><table width="707" height="70" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td width="707" height="70"><form name="form_reg" method="post" action="cart_checkout.asp" onSubmit="return chkreginfo(form_reg,'all')">


<table width="750" height="410" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td height="410">
		  <table width="680" border="0" align="center" cellpadding="0" cellspacing="1">
		    <tr>
              <td>
 <table width="680" border="0" align="center" cellpadding="0" cellspacing="0">
			    <tr>
                  <td height="40"><table width="80" height="22" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td bgcolor="#CCCCCC"><div align="center">�ջ�����Ϣ</div></td>
                    </tr>
                  </table></td>
                  <td colspan="3">&nbsp;<font color="#FF0000">*</font>&nbsp;<font color="#999999">�������ȷ��д���ĸ�����ϸ��Ϣ��</font></td>
                  </tr>
                <tr>
                  <td width="120" height="30"><div align="right">�ջ��ˣ�</div></td>
                  <td colspan="3">&nbsp;<input type="text" name="recuser" size="20" class="inputcss" onBlur="chkreginfo(form_reg,0)"><div id="chknew_recuser" style="color:#FF0000"></div></td>
                </tr>
				
                <tr>
                  <td height="30"><div align="right">��ϸ��ϵ��ַ��</div></td>
                  <td height="30" colspan="3">&nbsp;<input type="text" name="address" size="60" class="inputcss" onBlur="chkreginfo(form_reg,1)"><div id="chknew_address" style="color:#FF0000"></div></td>
                </tr>
                <tr>
                  <td height="30"><div align="right">�������룺</div></td>
                  <td height="30" colspan="3">&nbsp;<input type="text" name="yb" size="20" class="inputcss" onBlur="chkreginfo(form_reg,2)"><div id="chknew_yb" style="color:#FF0000"></div></td>
                </tr>
				<tr>
                  <td height="30"><div align="right">QQ���룺</div></td>
                  <td height="30" colspan="3">&nbsp;<input type="text" name="qq" size="20" class="inputcss" onBlur="chkreginfo(form_reg,3)"><div id="chknew_qq" style="color:#FF0000"></div></td>
                </tr>
				<tr>
                  <td height="30"><div align="right">E-mail��</div></td>
                  <td height="30" colspan="3">&nbsp;<input type="text" name="email" size="20" class="inputcss" onBlur="chkreginfo(form_reg,4)"><div id="chknew_email" style="color:#FF0000"></div></td>
                </tr>
                <tr>
                  <td height="30">&nbsp;</td>
                  <td height="30" colspan="3">&nbsp;<font color="#FF0000">*</font>&nbsp;<font color="#999999">�������ȷ��д������ϵ��ַ���ʱ࣬��ȷ�������ͻ���˳���ﵽ��</font></td>
                  </tr>
                <tr>
                  <td height="30"><div align="right">�ƶ��绰��</div></td>
                  <td width="150" height="30">&nbsp;<input type="text" name="mtel" size="20" class="inputcss" onBlur="chkreginfo(form_reg,5)"><div id="chknew_mtel" style="color:#FF0000"></div></td>
                  <td width="67"><div align="right">�̶��绰��</div></td>
                  <td width="343">&nbsp;<input type="text" name="gtel" size="20" class="inputcss" onBlur="chkreginfo(form_reg,6)"><div id="chknew_gtel" style="color:#FF0000"></div></td>
                </tr>
                
               
				
				
              </table></td>
            </tr>
          </table><br>
		  <table width="680" border="0" align="center" cellpadding="0" cellspacing="1">
  <tr>
    <td height="78"><table width="680" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td width="120" height="25"><table width="80" height="22" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td bgcolor="#CCCCCC"><div align="center">�ʵݷ�ʽ</div></td>
          </tr>
        </table></td>
        <td width="560">&nbsp;<font color="#FF0000">*</font>&nbsp;<font color="#999999">��ѡ���ͻ���ʽ��</font></td>
      </tr>
      
      <tr>
        <td height="30">&nbsp;</td>
        <td height="30">
        
          <input type="radio" name="shfs" value="1" checked="checked"/>
          ��ͨ�ʵ�<br><br>
          <input type="radio" name="shfs" value="2" />
         �����ؿ�ר��&nbsp;EMS        </td>
      </tr>
    </table></td>
  </tr>
</table>

		  
		 <table width="680" height="40" border="0" align="center" cellpadding="0" cellspacing="1">
  <tr>
    <td><table width="680" height="27" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td width="120" height="27"><table width="80" height="22" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td bgcolor="#CCCCCC"><div align="center">���ɶ���</div></td>
          </tr>
        </table></td>
        <td width="558"><input name="image" type="image"  src="images/sy_15.jpg">&nbsp;&nbsp;<img src="images/sy_17.jpg" width="60" height="25" onClick="form_reg.reset()" style="cursor:hand"/></td>
      </tr>
    </table></td>
    </tr>
</table>



</td>
        </tr>
      </table></form>
</td>
      </tr>
    </table>
	<%end if%>
	</td>
    <td width="41">&nbsp;</td>
  </tr>
  <tr>
    <td height="13" colspan="3">&nbsp;</td>
  </tr>
  <tr>
    <td height="50" colspan="3">&nbsp;</td>
  </tr>
</table>
</body>
</html>