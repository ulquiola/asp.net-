<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="style.css" rel="stylesheet">
<style type="text/css">
<!--
body {
	margin:0px;
	background-color:#b8dbf9;
}
-->
</style>
<script language="javascript">
function enterkey(){
	if(event.keyCode == 116){				//��������ǡ�F5����
		alert('��ֹˢ��');				//���������
		event.keyCode = 0;				//����������
		return false;
	}
}
document.onkeydown=enterkey;				//����������onkeydown�¼�
window.onload = function(){
	document.getElementById("message").focus();

}
</script>
<body  oncontextmenu=self.event.returnValue=false>
<table width="100%" height="115"  border="0" cellpadding="0" cellspacing="0">
<form action="main.asp" method="post" name="form1" target="mainFrame">
	<tr>
      <td width="75%" height="35">&nbsp;[<%=session("UserName")%>]
          <select name="face" >
            <option  value="�ޱ����">�ޱ����</option>
            <option value="΢Ц��" selected>΢Ц��</option>
            <option value="Ц�Ǻǵ�">Ц�Ǻǵ�</option>
            <option value="�����">�����</option>
            <option value="�����">�����</option>
            <option value="������">������</option>
            <option value="�Ҹ���">�Ҹ���</option>
            <option value="�����">�����</option>
            <option value="����ӯ����">����ӯ����</option>
            <option value="���������">���������</option>
            <option value="�����">�����</option>
            <option value="���������">���������</option>
            <option value="��ݺݵ�">��ݺݵ�</option>
            <option value="������">������</option>
            <option value="������">������</option>
            <option value="�����ֻ���">�����ֻ���</option>
            <option value="ͬ���">ͬ���</option>
            <option value="�ź���">�ź���</option>
            <option value="������Ȼ��">������Ȼ��</option>
            <option value="�����">�����</option>
            <option value="����˹���">����˹���</option>
            <option value="�޾���ɵ�">�޾���ɵ�</option>
          </select>
      ��
	  <%
		  if Request.QueryString("name") <>"" then
			toname=Request.QueryString("name")
		  else 
			  toname="������"
		  end if
	  %>
      <input name="to" type="text" class="Sytle_text" id="to" value="<%=toname%>" size="20" readonly="readonly" style=" border-color: 8e999f">
      ˵      
      <input name="from" type="hidden" id="from" value="<%=session("UserName")%>">
	  </td>
      <td width="25%">&nbsp;������ɫ��
          <select name="color" size="1" id="select">
            <option selected>Ĭ����ɫ</option>
            <option style="color:#FF0000" value="FF0000">��ɫ����</option>
            <option style="color:#0000FF" value="0000ff">��ɫ����</option>
            <option style="color:#ff00ff" value="ff00ff">��ɫ����</option>
            <option style="color:#009900" value="009900">��ɫ�ഺ</option>
            <option style="color:#009999" value="009999">��ɫ��ˬ</option>
            <option style="color:#990099" value="990099">��ɫ�н�</option>
            <option style="color:#990000" value="990000">��ҹ�˷�</option>
            <option style="color:#000099" value="000099">��������</option>
            <option style="color:#999900" value="999900">�����Ʒ�</option>
            <option style="color:#ff9900" value="ff9900">�ֽ�����</option>
            <option style="color:#0099ff" value="0099ff">��������</option>
            <option style="color:#9900ff" value="9900ff">��������</option>
            <option style="color:#ff0099" value="ff0099">���İ�ʾ</option>
            <option style="color:#006600" value="006600">ī�����</option>
            <option style="color:#999999" value="999999">��������</option>
        </select></td>
    </tr>
	<script language="javascript">
	function check(){
	if(form1.message.value=="")
	{alert("�������ݲ���Ϊ�գ�");form1.message.focus();return;}
	form1.submit();  //�ύ��
	form1.reset();  //���ñ�
	form1.message.focus();
	}
	</script>
<script language="javascript">
	function Exit1(){
	parent.location.href="exit.asp";
	}
	</script>
    <tr>
      <td height="35">&nbsp;
          <input name="message" type="text" class="Sytle_text" id="message" onKeyDown="Javascript:if(event.keyCode==13){check()}" size="60" style="border-color: #8e999f;">
          <input name="send" type="button" id="send" value="&nbsp;" onClick="check()" style="background-image:url(images/send.gif); background-position:center; background-repeat:no-repeat; width:42px; height:21px; border: 1px #C67131 solid;"></td>
      <td><div align="center">
          <input type="button" name="Submit2" value="&nbsp;" onClick="Exit1()" style=" background-image:url(images/quit.gif); background-position:center; background-repeat:no-repeat; width:104px; height:23px; border: 1px #C67131 solid;">
      </div></td>
    </tr>
</form>	
    <tr>
      <td height="35" colspan="2">
        <table>
	  <tr>
	  <td>
<form action="Upload_OK.asp" method="post" enctype="multipart/form-data" name="form2" target="mainFrame">
	  ��ѡ��ͼƬ:
	  <input name="filename" type="file" id="filename" size="40" class="Sytle_text" style=" border-color: #8e999f;">
	  <input type="button" name="SendImg" value="&nbsp;" onClick="JavaScript:form2.submit();" style=" background-image:url(images/sendpic.gif); background-position:center; background-repeat:no-repeat; width:61px; height:21px; border: 1px #C67131 solid;">
</form>	  
	  </td>
	  </tr>
	  </table>
	</td>
    </tr>
  </table>
  </body>