<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�ޱ����ĵ�</title>
<style type="text/css">
<!--
.STYLE15 {
	color: #FFFFFF;
	font-size: 9pt;
}
body {
	margin-top: 0px;
}
.STYLE17 {
	font-size: 9pt;
	color: #000000;
}
.STYLE18 {
	font-size: 10pt;
	font-weight: bold;
}
-->
</style>
</head>
<body>
<table width=380 border=1 align=center cellpadding=0 cellspacing=0 bordercolor=#000000>
  <tr>
    <td width="376" height="19" valign=middle bordercolor=#FFFFFF bgcolor=#A90303>    <div align="center" class="STYLE15 STYLE18">�� �� �� Ϣ �� ѯ</div></td>
  </tr>
  <tr>
    <td valign=top bordercolor=#FFFFFF bgcolor=#FFFFFF>&nbsp;</td>
  </tr>
  <tr>
    <td height="89" valign=top bordercolor=#FFFFFF bgcolor=#FFFFFF>
      <div align=center>
	  <script language="javascript">
	  	function get()							//�����Զ��庯��
			{
				if(form1.OrderID.value=="")		//�ж϶������Ƿ�Ϊ��
					{alert("�����붩������");	//������ʾ�Ի���
					form1.OrderID.focus();		//��ȡ����
					}
				if(isNaN(form1.OrderID.value))	//�ж϶������Ƿ���������
				{
				alert("�����������Ϊ���ֵ�?");	//������ʾ�Ի���
				form1.OrderID.focus();			//��ȡ����
				}
				else
					{
					window.close();				//�رմ���
					window.open("order_index.asp?OrderID="+form1.OrderID.value);	//��ָ����ҳ��
					}
			}			
	  </script>
        <form action="" method="post" name="form1">
          <span class="STYLE17">�����붩�����룺</span>
          <input name="OrderID" type="text" id="OrderID">
          <br>
            <br>
            <input name="Submit" type="submit" class="go-wenbenkuang" value="��ѯ" onClick="get()">            
         	&nbsp;&nbsp;&nbsp;&nbsp;
	 	  <input name="Submit2" type="button" class="go-wenbenkuang" value="����" onclick="javascript:window.close()"> 
            <br>
        </form>
      </div>
      <div align="right"></div>
    <div align="right"></div></td>
  </tr>
</table>
</body>
</html>
