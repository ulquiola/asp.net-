<!--#include file="Conn/conn.asp" -->
<!--#include file="js/Function.asp" -->
<%
Set rs = Server.CreateObject("ADODB.RecordSet")
sql="select * from tb_user"
rs.open sql,conn,1,3
%>
<HTML>
<HEAD>
<title>��Ӻ���</title>
<script language="javascript">
function mycheckadd()
{//�ڽ��к������ʱ������Ҫ�ж�ѡ���ı������Ƿ����û�
if(form1.sel_place2.length<1){
	alert("��ѡ��Ҫ��ӵĺ���!")			//������ʾ�Ի���
	return 
}
var value="";								//Ϊvalue��ֵ
for (i=0;i<form1.sel_place2.length;i++){	
	value=value+form1.sel_place2.options[i].value+",";
	//Ϊvalue��ֵ
}
value=value.substr(0,value.length-1);
//����ӵĶ���û���ͷ����Ϣ��ֵ���������ʽ���ݵ�����ҳ��
//����value��ʾ���ݵ�����Ҫ��ӵ��û�IDֵ
//GroupID���ڶ�̬���ݵ�IDֵ
window.opener.frames("mainfrm").frames("right").location = "Addfriend_deal.asp?ID=" + value + "&GroupID=<%=request.querystring("GroupID")%>"
self.close() 
}
</script>
<script language="javascript">
function allsel(n1,n2)
{//ʵ���û���Ϣ��������Ƴ�
  while(n1.selectedIndex!=-1)
  {
  	var indx=n1.selectedIndex;						//��������������ֵ
  	var t=n1.options[indx].text;					//��������������ֵ
	t1=n1.options[indx].value;						//Ϊt1������ֵ
  	n2.options.add(new Option(t,t1));				
  	n1.remove(indx);								
  }
}
</script>
<script language="javascript"> 
function mychecktext()
//���ư�ť�ǲ�����
{ 
if(document.getElementById("content").value.length>0){ 
document.getElementById("chazhao").disabled=""; 
//���ð�ťΪ����״̬
}else{ 
document.getElementById("chazhao").disabled="disabled"; 
//���ð�ťΪ������״̬
} 
} 
//��ѯʱ������ѯ�����������ɫѡ��
function sss(myForm){
	for(i=0;i<myForm.Userlist.length;++i){
		if(myForm.Userlist.options[i].value==myForm.content.value)
		//�ж�ѡ�е�ֵ�ͱ��������Ƿ����
		{
			myForm.Userlist.options[i].selected=true;
			//����ǰ�����ݴ���ѡ��״̬
		}
		
	}
}
</script> 
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"><style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style></HEAD>
<style type="text/css">
<!--
.STYLE3 {font-size: 9pt; color: #000000; }
.STYLE6 {color: #FF0000}
-->
</style>
<BODY style="border:none">
<form name="form1">
<table width="476" height="324" border="0" cellpadding="0" cellspacing="0" background="Images/bg.gif" id="AutoNumber1">
  <tr>
    <td colspan="3"><span class="STYLE3">&nbsp;&nbsp;&nbsp;&nbsp;<br>
      &nbsp;&nbsp;&nbsp;&nbsp;�������û�ID����б���ѡ��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;ѡ�е��û��� </span></td>
  </tr>
  <tr>
    <td width="46%" height="281" valign="top"><br>
      <table width="221" height="26" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td width="24">&nbsp;</td>
        <td width="147"><input name="content" type="text"  id="content" size="18" maxlength="12" height="12 " onKeyUp="mychecktext()"></td>
        <td width="50"><input type="button" name="chazhao" value="����" disabled="disabled" id="chazhao" onClick="sss(form1)"></td>
      </tr>
    </table>      
    &nbsp;&nbsp;  
    <select name="Userlist" size="10" multiple style="width:186px ">
            <%while not rs.EOF %>
			<%
			if Session("username1")<>rs("Username") then
			%>
            <option value="<%=rs("ID")%>"><%=rs("id")%>&nbsp;(<%=rs("Username")%>)</option>
			<%
			end if
    		rs.Movenext() 
    		wend
	%>
        </select>
          <br><br>
        </p>
        <p align="right"><br>
          <br>
&nbsp;&nbsp;&nbsp;<br>        
          <br>
      </p>      </td>
    <td width="14%" valign="top"><p align="center">&nbsp;</p>
      <p align="center">&nbsp;</p>
      <p>
        <input type="button" name="sure2" value="����&gt;&gt;" onClick="allsel(document.form1.Userlist,document.form1.sel_place2);">
      </p>
      <p>
        <input type="button" name="sure1" value="&lt;&lt;�Ƴ�" onClick="allsel(document.form1.sel_place2,document.form1.Userlist);">
      </p></td>
    <td width="40%" valign="top">
      <p><br>
          <select name="sel_place2" size="12" multiple class="wenbenkuang" id="sel_place2" style="width:178px ">
          </select>
          <br>
 </p><br><br>
      <p>         
        <input type="button" value="  ���  " name="B32" onClick="mycheckadd()">
        &nbsp;
        <input type="button" value="  ȡ��  " name="B3" onClick="window.close()">
&nbsp;      </p></td>
  </tr>
</table>
</form>
</BODY>
</HTML>
