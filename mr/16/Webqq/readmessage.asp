<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/conn.asp" -->
<!--#include file="js/Function.asp" -->
<%
Set rs = Server.CreateObject("ADODB.Recordset")			'������¼��
sql="SELECT * FROM tb_user where ID="&request.QueryString("UserID")&""	'��ѯ����
rs.open sql,conn,1,3													'�򿪼�¼��
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
<script language="javascript">
	window.resizeTo("410","211")					
	estate = "close"								//Ϊestate��������Ĭ��ֵ
	function talknote()
	{												//�����Զ��庯��
	     if(estate=="close"){						//�ж�estate������ֵ�Ƿ�Ϊclose
			window.resizeTo("410","411")
			document.all.dibu.style.display=""		//��ʾ�ظ���Ϣҳ��
			window.frames("bottom").location = "talknote.asp?UserID=<%=request.QueryString("UserID")%>"
			//��ת��ָ����ASPҳ���У�����ʾ������Ϣ
			estate = "open"							//Ϊestate��������Ĭ��ֵ
	     }
	     else{
			window.resizeTo("410","211")
			document.all.dibu.style.display="none"
			estate = "close"
	     }
	}
	function mycheck1(){
	  //window.event.returnValue = false;
	}
	function formSubmit(){
	    keyvalue = window.event.keyCode
		if(keyvalue==13){
            window.location="message_send.asp?userID=<%=request.querystring("UserID")%>"
		}
		else{
			return
		}
	}
	
	function closed(){
	     window.moveTo(2000,200)
	}
</script>
<style type="text/css">
<!--
.Fieldset {  font-size: 12px; border: 1px inset; padding-top: 2px; padding-right: 4px; padding-bottom: 2px; padding-left: 4px; margin-top: 2px; margin-right: 4px; margin-bottom: 2px; margin-left: 4px}
.button {  font-size: 12px; text-align: center; vertical-align: middle}
td {  padding-right: 2px; padding-left: 2px}
.text {  background-color: #ECE9D8; height: 18px}
.STYLE1 {
	font-size: 9pt;
	color: #000000;
}
.STYLE2 {font-size: 9pt}
-->
</style>
</head>
<body scroll="no" topmargin="0" leftmargin="0" style="border:none" bgcolor="#ECE9D8" id="bodyID" oncontextmenu="mycheck1()" onkeydown="formSubmit()">
<table border="0" cellpadding="0" cellspacing="0" bordercolor="#111111" width="100%" id="AutoNumber1" height="100%" style="border-collapse: collapse">
  <form name="form1" action="messageSend_deal.asp" method="post">
  <tr height="50">
    <td>
     <fieldset class="Fieldset"><legend><span lang="zh-cn">����</span></legend> 
      <table border="1" cellpadding="0" cellspacing="0" width="100%">
        <tr> 
          <td width="25%">
            <div align="right" class="STYLE1">�û���:</div>          </td>
          <td width="25%" class="STYLE2">
             <a href="#" onClick="javascript:window.open('chakan.asp?id=<%=rs("ID")%>&type=new','','width=350,height=250,toolbar=no,location=no,status=no,menubar=no')"><%=rs("UserName")%></a></td>
          <td width="25%">
            <div align="right" class="STYLE2">�ǳ�:</div>          </td>
          <td width="25%">
            <span class="STYLE2"><%=rs("state")%>            </span>
            <input type="hidden" value="<%=rs("ID")%>" name="ToUserID">
			<input type="hidden" value="<%=request.Cookies("fusernickname")%>" name="fusernickname">
			<input type="hidden" value="<%=request.Cookies("fuserid")%>" name="fuserid">		  </td>
        </tr>
      </table>
      </fieldset>
    </td>
  </tr>
  <tr id="topTr">
    <td width="100%"><iframe name="frmTop" style="width:100%;height:100%" src="message_list.asp?UserID=<%=request.QueryString("UserID")%>"></iframe></td>
  </tr>
  <tr height="28" id="barTr">
    <td width="100%">
    <table width="100%" border="0" cellpadding="0" cellspacing="0">
		<tr>
          <td>
           <input type="button" value="����Ϣ" name="B3" class="button" onClick="window.location='message_send.asp?userID=<%=request.querystring("UserID")%>'"> 
          </td>
			<td align="right">
			<input type="button" value="�����¼" name="B3" onClick="talknote()" class="button">
			<input type="button" value="�ر�" name="B3" class="button" onClick="closed()">
          </td>
		</tr>
    </table>
    </td>
  </tr>
  <tr id="dibu" height="200" style="display:none">
    <td width="100%"><iframe name="bottom" style="width:100%;height:100%" src=""></iframe></td>
  </tr>
  </form>
</table>
</body>
</html>