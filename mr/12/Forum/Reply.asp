<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/Conn.asp"-->
<!--�������ݿ������ļ�-->
<%
Set rs_U=Server.CreateObject("ADODB.RecordSet")							'������¼��
SQL="Select * From tb_User Where UserName='"&Session("UserName")&"'"	'��ѯ����
rs_U.Open SQL,Conn,1,3													'�򿪼�¼��
If (rs_U.Eof and rs_U.Bof) or Session("UserName")="Friend" Then			'�жϵ�ǰ�û��Ƿ��¼%>		
<script language="javascript">
	alert("�Բ�����û��ע����¼���ܻظ���Ϣ��������Ϣ��");			//������ʾ�Ի���
	window.location.href="Login.asp?ReplyF=yes";						//��ת��ָ����ҳ��
</script>
<%
Response.End()														
Else
	Set rs_U=Server.CreateObject("ADODB.RecordSet")							'������¼��
	SQL="Select * From tb_User Where UserName='"&Session("UserName")&"'"	'��ѯ����
	rs_U.Open SQL,Conn,1,3													'�򿪼�¼��
	UID=rs_U("ID")															'ΪUID������ֵ
	UserName=session("UserName")											'ΪUserName������ֵ
	Email=rs_U("Email")														'ΪEmail������ֵ
	homepage=rs_U("homepage")												'Ϊhomepage������ֵ
	ICO=rs_U("ICO")															'ΪICO������ֵ
	QQ=rs_U("QQ")															'ΪQQ������ֵ
	IP=Request.ServerVariables("REMOTE_ADDR")								'Ӧ��ServerVariables������ȡ�û���IP��ַ
End If
%>
<%
If Request.QueryString("ReplyID")<>"" and Request.QueryString("state")="0" Then		'�ظ���Ϣ��ID����Ϊ��
	ReplyID=Request.QueryString("ReplyID")											'��ȡ�ظ���Ϣ��ID
	Set rs1=Server.CreateObject("ADODB.RecordSet")									'������¼��
	SQL="Select * From tb_Reply Where ID="&ReplyID									'��ѯ����
	rs1.Open SQL,Conn,1,3															'�򿪼�¼��
	if not rs1.eof and not rs1.bof then												'��ѯ�Ƿ��м�¼��Ϣ
			if Request.QueryString("state")="0" then								
			i=request.QueryString("i")												'Ϊi������ֵ
			title="���ԣ�"&i&"¥"													'Ϊtitle������ֵ
			content=rs1("content")													'Ϊcontent������ֵ
			tttt="[FIELDSET][LEGEND]"&title&"[/LEGEND]"&chr(13)&content&chr(13)&"[/FIELDSET]"&chr(13)&chr(13)&"�ظ���"&chr(13)	'Ӧ��FIELDSET��ǩ�����õ���ʽ��ʾ��ǰ���õ��Ǽ�¥�û�����Ϣ
			end if
	end if
	rs1.close()																		'�رռ�¼��
	Set rs1=Nothing																	'�ͷ��ڴ�ռ�
End If
%>
<%
If Request.QueryString("TopicID")<>"" then 											'�ж�������Ϣ��ID�Ƿ�Ϊ��
	TopicID=Request.QueryString("TopicID")											'ΪTopicID������ֵ
end if
If TopicID<>"" and Request.QueryString("state")="1" Then							'�ж�����IDֵ�Ƿ�ղ��ж��Ƿ���¥���������Ϣ
	Set rs=Server.CreateObject("ADODB.RecordSet")									'������¼��
	SQL="Select * From view_board Where ID="&Request.QueryString("TopicID")			'��ѯ������Ϣ
	rs.Open SQL,Conn,1,3															'�򿪼�¼��
	if not rs.eof and not rs.bof then
		if Request.QueryString("state")="1" then									'�ж�stateֵ�Ƿ�Ϊ1
			title="���ԣ�¥��"														'Ϊtitle������ֵ
			content=rs("content")													'Ϊcontent������ֵ
			tttt="[FIELDSET][LEGEND]"&title&"[/LEGEND]"&chr(13)&content&chr(13)&"[/FIELDSET]"&chr(13)&chr(13)&"�ظ���"&chr(13)
			'Ӧ��FIELDSET��ǩ�����õ���ʽ��ʾ��ǰ���õ�¥����Ϣ
		end if
	end if
	rs.close()																		'�رռ�¼��
	Set rs=Nothing																	'�ͷ��ڴ�ռ�
end if
%>
<html>
<head>
<title>�ظ�������Ϣ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Css/style.css" rel="stylesheet">
<script src="JS/UBBCode.asp"></script>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style></head>
<script language="javascript">
function check(){
	if(form1.subject.value==""){
		alert("�ظ����ⲻ����Ϊ�գ�");form1.subject.focus();return;
	}
	if(form1.content.value==""){
		alert("�ظ����ݲ�����Ϊ�գ�");form1.content.focus();return;
	}
	form1.submit();
}
</script>
<SCRIPT language=JavaScript>
<!--
var LastCount =0;
function CountStrByte(Message,Total,Used,Remain){ //�ֽ�ͳ��
 var ByteCount = 0;
 var StrValue  = Message.value;
 var StrLength = Message.value.length;
 var MaxValue  = Total.value;

 if(LastCount != StrLength) { // �ڴ��жϣ�����ѭ������
	for (i=0;i<StrLength;i++){
		ByteCount  = (StrValue.charCodeAt(i)<=256) ? ByteCount + 1 : ByteCount + 2;
      if (ByteCount>MaxValue) {
      	Message.value = StrValue.substring(0,i);
			alert("�ظ�������಻�ܳ��� " +MaxValue+ " ���ֽڣ�\nע�⣺һ������Ϊ���ֽڡ�");
         ByteCount = MaxValue;
         break;
      }
	}
   Used.value = ByteCount;
   Remain.value = MaxValue - ByteCount;
   LastCount = StrLength;
 }
}

//-->
</SCRIPT>
<%
'ת��UBB����
Function UBBCodeDeal(InString)
	InString1=Server.HTMLEncode(InString)
	InString1=Replace(InString1,"[","<")
	InString1=Replace(InString1,chr(13),"<BR>")
	UBBCodeDeal=Replace(InString1,"]",">")
End Function
%>
<body> 
<table width="777" height="341"  border="0" align="center" cellpadding="-2" cellspacing="-2" bgcolor="#FFFFFF">
  <tr>  <td valign="top"> 
   <table width="100%" border="0" cellpadding="-2" cellspacing="-2" class="tableBorder"> 
    <tr> 
       <td></td> 
     </tr> 
    <tr> 
       <td></td> 
     </tr> 
    <tr> 
       <td height="2"></td> 
     </tr> 
    <tr> 
       <td valign="top">
  <table width="100%" height="205"  border="0" cellpadding="-2" cellspacing="-2" class="tableBorder"> 
    <tr> 
      <td width="13" height="32" background="Images/bg.gif">&nbsp;</td> 
      <td width="689" background="Images/bg.gif" style="color:#FFFFFF"> �� �ظ�������Ϣ �� </td> 
      <td width="73" background="Images/bg.gif">&nbsp;</td> 
    </tr> 
    <tr valign="top"> 
     <td height="171" colspan="3"> 
     <table width="100%" height="129"  border="0" cellpadding="-2" cellspacing="-2"> <tr> 
       <td valign="top"> 
        <table width="100%" height="265"  border="0" cellpadding="-2" cellspacing="-2" class="tableBorder_Top"> 
         <tr> 
            <td height="5"></td> 
          </tr> 
         <tr> 
          <td width="280" height="263" valign="top"><div align="center"> 
              <p> === �ظ�����Ϣ ===<br> 
              <%=UserName%></p> 
              <p><img src="Images/head/<%=ICO%>" width="60" height="60" border=0> </p> 
              <p>���ǣ�<%=rs_U("sex")%>��</p> 
              <p> <img src="Images/email.gif" alt="<%=UserName%>��Email�ǣ�" width="45" height="16"> Email��<%=Email%></p> 
              <p> <img src="Images/oicq.gif" alt="<%=UserName%>��QQ�����ǣ�" width="16" height="16"> QQ:<%=QQ%></p> 
              <p> <img src="Images/ip.gif" alt="IPΪ��" width="16" height="15"> IP��<%=IP%></p> 
            </div></td> 
         <td width="24" background="Images/line.gif"></td> 
         <td width="686" valign="top"> 
          <form action="Reply_deal.asp" method="post" name="form1"> 
            <table width="100%"  border="0" cellspacing="-2" cellpadding="-2"> 
              <tr> 
                <td width="14%" height="50" align="center">��ǰ�û���</td> 
                <td colspan="2" width="86%" style="padding-left:10px"><%=session("username")%></td> 
              </tr> 
              <tr> 
                <td height="20" colspan="2" align="center" style="padding-left:10px"><hr size="1" noshade></td> 
              </tr> 
              <tr> 
                <td height="40" align="center" style="padding-left:10px">�ظ����⣺</td> 
                <td height="40" style="padding-left:10px"><input name="subject" type="text" id="subject" size="60"></td> 
              </tr> 
              <tr> 
                <td height="145" align="center">�ظ����ݣ�</td> 
                <td><textarea name="content" cols="70" rows="14" class="wenbenkuang" id="content"
					   onkeydown=CountStrByte(this.form.content,this.form.total,this.form.used,this.form.remain);
					    onkeyup=CountStrByte(this.form.content,this.form.total,this.form.used,this.form.remain);><%=tttt%></textarea>
                  <br></td> 
              </tr> 
              <tr> 
                <td height="33" align="center" style="padding-left:10px"><p>�ֽڣ�</p></td> 
                <td style="padding-left:10px" class="word_grey">�������
                  <input name="total" type="text" disabled class="noborder" id="total"  value="1600" size="4"> 
                  ���ֽ� �����ֽڣ�&nbsp; 
                  <input name="used" type="text" disabled class="noborder" id="used"  value="0" size="4"> 
                  ʣ���ֽڣ�
                  <input name="remain" type="text" disabled class="noborder" id="remain" value="1600" size="4"> </td> 
              </tr> 
              <tr> 
                <td height="34" align="center" style="padding-left:10px">&nbsp;</td> 
                <td class="word_grey"><input name="Button" type="button" class="btn_grey" value="�ύ�ظ�" onClick="check()"> 
                  <input name="Submit2" type="reset" class="btn_grey" value="��д�ظ�"> 
				 <input name="HIP" type="hidden" id="HIP" value="<%=Request.ServerVariables("REMOTE_ADDR")%>"> 
                <input name="HTopicID" type="hidden" value="<%=TopicID%>"> 
                <input name="UID" type="hidden" id="UID" value="<%=UID%>"> 
              </td> 
              </tr> 
               </table> 
          </form> 
         </td> 
          </tr> 
         <tr> 
            <td height="5"></td> 
          </tr> 
       </table> 
      </td> 
       </tr> 
        </table> 
    </td> 
     </tr> 
      </table> 
  </td> 
  </tr> 
</table> 
</body>
</html>

