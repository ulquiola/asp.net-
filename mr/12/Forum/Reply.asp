<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/Conn.asp"-->
<!--包含数据库连接文件-->
<%
Set rs_U=Server.CreateObject("ADODB.RecordSet")							'创建记录集
SQL="Select * From tb_User Where UserName='"&Session("UserName")&"'"	'查询数据
rs_U.Open SQL,Conn,1,3													'打开记录集
If (rs_U.Eof and rs_U.Bof) or Session("UserName")="Friend" Then			'判断当前用户是否登录%>		
<script language="javascript">
	alert("对不起！您没有注册或登录不能回复信息或引用信息！");			//弹出提示对话框
	window.location.href="Login.asp?ReplyF=yes";						//跳转到指定的页面
</script>
<%
Response.End()														
Else
	Set rs_U=Server.CreateObject("ADODB.RecordSet")							'创建记录集
	SQL="Select * From tb_User Where UserName='"&Session("UserName")&"'"	'查询数据
	rs_U.Open SQL,Conn,1,3													'打开记录集
	UID=rs_U("ID")															'为UID变量赋值
	UserName=session("UserName")											'为UserName变量赋值
	Email=rs_U("Email")														'为Email变量赋值
	homepage=rs_U("homepage")												'为homepage变量赋值
	ICO=rs_U("ICO")															'为ICO变量赋值
	QQ=rs_U("QQ")															'为QQ变量赋值
	IP=Request.ServerVariables("REMOTE_ADDR")								'应用ServerVariables方法获取用户的IP地址
End If
%>
<%
If Request.QueryString("ReplyID")<>"" and Request.QueryString("state")="0" Then		'回复信息的ID不能为空
	ReplyID=Request.QueryString("ReplyID")											'获取回复信息的ID
	Set rs1=Server.CreateObject("ADODB.RecordSet")									'创建记录集
	SQL="Select * From tb_Reply Where ID="&ReplyID									'查询数据
	rs1.Open SQL,Conn,1,3															'打开记录集
	if not rs1.eof and not rs1.bof then												'查询是否有记录信息
			if Request.QueryString("state")="0" then								
			i=request.QueryString("i")												'为i变量赋值
			title="引自："&i&"楼"													'为title变量赋值
			content=rs1("content")													'为content变量赋值
			tttt="[FIELDSET][LEGEND]"&title&"[/LEGEND]"&chr(13)&content&chr(13)&"[/FIELDSET]"&chr(13)&chr(13)&"回复："&chr(13)	'应用FIELDSET标签以引用的形式显示当前引用的是几楼用户的信息
			end if
	end if
	rs1.close()																		'关闭记录集
	Set rs1=Nothing																	'释放内存空间
End If
%>
<%
If Request.QueryString("TopicID")<>"" then 											'判断主题信息的ID是否为空
	TopicID=Request.QueryString("TopicID")											'为TopicID变量赋值
end if
If TopicID<>"" and Request.QueryString("state")="1" Then							'判断主题ID值是否空并判断是否是楼主发表的信息
	Set rs=Server.CreateObject("ADODB.RecordSet")									'创建记录集
	SQL="Select * From view_board Where ID="&Request.QueryString("TopicID")			'查询数据信息
	rs.Open SQL,Conn,1,3															'打开记录集
	if not rs.eof and not rs.bof then
		if Request.QueryString("state")="1" then									'判断state值是否为1
			title="引自：楼主"														'为title变量赋值
			content=rs("content")													'为content变量赋值
			tttt="[FIELDSET][LEGEND]"&title&"[/LEGEND]"&chr(13)&content&chr(13)&"[/FIELDSET]"&chr(13)&chr(13)&"回复："&chr(13)
			'应用FIELDSET标签以引用的形式显示当前引用的楼主信息
		end if
	end if
	rs.close()																		'关闭记录集
	Set rs=Nothing																	'释放内存空间
end if
%>
<html>
<head>
<title>回复主题信息！</title>
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
		alert("回复主题不允许为空！");form1.subject.focus();return;
	}
	if(form1.content.value==""){
		alert("回复内容不允许为空！");form1.content.focus();return;
	}
	form1.submit();
}
</script>
<SCRIPT language=JavaScript>
<!--
var LastCount =0;
function CountStrByte(Message,Total,Used,Remain){ //字节统计
 var ByteCount = 0;
 var StrValue  = Message.value;
 var StrLength = Message.value.length;
 var MaxValue  = Total.value;

 if(LastCount != StrLength) { // 在此判断，减少循环次数
	for (i=0;i<StrLength;i++){
		ByteCount  = (StrValue.charCodeAt(i)<=256) ? ByteCount + 1 : ByteCount + 2;
      if (ByteCount>MaxValue) {
      	Message.value = StrValue.substring(0,i);
			alert("回复内容最多不能超过 " +MaxValue+ " 个字节！\n注意：一个汉字为两字节。");
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
'转换UBB代码
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
      <td width="689" background="Images/bg.gif" style="color:#FFFFFF"> ≡ 回复主题信息 ≡ </td> 
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
              <p> === 回复人信息 ===<br> 
              <%=UserName%></p> 
              <p><img src="Images/head/<%=ICO%>" width="60" height="60" border=0> </p> 
              <p>我是：<%=rs_U("sex")%>生</p> 
              <p> <img src="Images/email.gif" alt="<%=UserName%>的Email是：" width="45" height="16"> Email：<%=Email%></p> 
              <p> <img src="Images/oicq.gif" alt="<%=UserName%>的QQ号码是：" width="16" height="16"> QQ:<%=QQ%></p> 
              <p> <img src="Images/ip.gif" alt="IP为：" width="16" height="15"> IP：<%=IP%></p> 
            </div></td> 
         <td width="24" background="Images/line.gif"></td> 
         <td width="686" valign="top"> 
          <form action="Reply_deal.asp" method="post" name="form1"> 
            <table width="100%"  border="0" cellspacing="-2" cellpadding="-2"> 
              <tr> 
                <td width="14%" height="50" align="center">当前用户：</td> 
                <td colspan="2" width="86%" style="padding-left:10px"><%=session("username")%></td> 
              </tr> 
              <tr> 
                <td height="20" colspan="2" align="center" style="padding-left:10px"><hr size="1" noshade></td> 
              </tr> 
              <tr> 
                <td height="40" align="center" style="padding-left:10px">回复主题：</td> 
                <td height="40" style="padding-left:10px"><input name="subject" type="text" id="subject" size="60"></td> 
              </tr> 
              <tr> 
                <td height="145" align="center">回复内容：</td> 
                <td><textarea name="content" cols="70" rows="14" class="wenbenkuang" id="content"
					   onkeydown=CountStrByte(this.form.content,this.form.total,this.form.used,this.form.remain);
					    onkeyup=CountStrByte(this.form.content,this.form.total,this.form.used,this.form.remain);><%=tttt%></textarea>
                  <br></td> 
              </tr> 
              <tr> 
                <td height="33" align="center" style="padding-left:10px"><p>字节：</p></td> 
                <td style="padding-left:10px" class="word_grey">最多允许
                  <input name="total" type="text" disabled class="noborder" id="total"  value="1600" size="4"> 
                  个字节 已用字节：&nbsp; 
                  <input name="used" type="text" disabled class="noborder" id="used"  value="0" size="4"> 
                  剩余字节：
                  <input name="remain" type="text" disabled class="noborder" id="remain" value="1600" size="4"> </td> 
              </tr> 
              <tr> 
                <td height="34" align="center" style="padding-left:10px">&nbsp;</td> 
                <td class="word_grey"><input name="Button" type="button" class="btn_grey" value="提交回复" onClick="check()"> 
                  <input name="Submit2" type="reset" class="btn_grey" value="重写回复"> 
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

