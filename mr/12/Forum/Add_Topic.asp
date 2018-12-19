<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/Conn.asp"-->
<%
	typeid=request.querystring("TypeID")				'获取表单元素ID的值并赋给ID变量
	set rs_T=Server.CreateObject("ADODB.RecordSet")		'创建记录集
	sql="Select * From tb_type"							'查询数据
	rs_T.open sql,conn,1,3								'打开记录集
	set rs_1=Server.CreateObject("ADODB.RecordSet")		'创建记录集
	sql1="Select * From view_smalltype where ID="&typeid'查询数据
	rs_1.open sql1,conn,1,3								'打开记录集

Set rs_T=Server.CreateObject("ADODB.RecordSet")			'创建记录集
SQL="Select * From tb_smalltype"						'查询数据
rs_T.Open SQL,Conn,1,3									'打开记录集
Set rs=Server.CreateObject("ADODB.RecordSet")			'创建记录集
SQL="Select * From tb_User Where UserName='"&Session("UserName")&"'"	'查询数据
rs.Open SQL,Conn,1,3													'打开记录集
If rs.Eof and rs.Bof Then												'判断是否有用户信息
%>
<script language="javascript">
	alert("先注册成为用户，才可以发表主题信息！");						//弹出提示对话框
	window.location.href="index.asp";									//跳转到指定的页面
</script>
<%
 Response.End()															'结束输出
 End If
 %>
<html>
<head>
<title>发表主题！</title>
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
function check()
{
	if(form1.subject.value==""){
		alert("主题信息不允许为空！");form1.subject.focus();return;
	}
	if(form1.content.value==""){
		alert("留言内容不允许为空！");form1.content.focus();return;
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
			alert("留言内容最多不能超过 " +MaxValue+ " 个字节！\n注意：一个汉字为两字节。");
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
<body>
<table width="777" height="341"  border="0" align="center" cellpadding="-2" cellspacing="-2" bgcolor="#FFFFFF">
  <tr>
    <td valign="top"><table width="100%" height="562" border="0" align="center" cellpadding="-2" cellspacing="-2" class="tableBorder">
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
          <td width="689" background="Images/bg.gif" style="color:#FFFFFF"> ≡ 发表主题


 ≡ </td>
          <td width="73" background="Images/bg.gif">&nbsp;</td>
        </tr>
        <tr valign="top">
          <td height="171" colspan="3"><table width="100%" height="129"  border="0" cellpadding="-2" cellspacing="-2">
            <tr>
              <td valign="top">
	  			  <table width="100%" height="265"  border="0" cellpadding="-2" cellspacing="-2" class="tableBorder_Top">
			  <tr>
			  <td height="5"></td>
			  </tr>
                <tr>
                  <td width="270" height="263" valign="top"><div align="center">
                    <p>
                       === 发帖人信息 ===<br>
                      <%=Session("UserName")%></p>
					  <%
					  UserName=rs("UserName")
					  Email=rs("Email")
					  homepage=rs("homepage")
					  ICO=rs("ICO")
					  QQ=rs("QQ")
					  %>
                    <p><img src="Images/head/<%=ICO%>" width="60" height="60" border=0> </p>
                    <p>我是：<%=rs("sex")%>生</p>
                    <p>
                        <img src="Images/email.gif" alt="<%=UserName%>的Email是：" width="45" height="16"> <%=Email%></p>
                    <p>
					    <img src="Images/oicq.gif" alt="<%=UserName%>的QQ号码是：" width="16" height="16"> QQ:<%=QQ%></p>
                    <p>
					      <img src="Images/ip.gif" alt="IP为：" width="16" height="15"> IP：<%=Request.ServerVariables("REMOTE_ADDR")%></p>
                  </div></td>
                  <td width="23" background="Images/line.gif"></td>
                  <td width="661" valign="top">
				  <form action="add_Topic_deal.asp" method="post" name="form1">
				  <table width="100%"  border="0" cellspacing="-2" cellpadding="-2">
                    <tr>
                      <td height="34" align="center" style="padding-left:10px">类别：</td>
                      <td class="word_grey">
<select name="TypeName">
<%For i=1 to rs_T.recordcount%>
<option value="<%=rs_T("ID")%>" <%if trim(rs_T("ID"))=typeid then %>selected <%end if%>><%=rs_T("smalltypename")%></option>
					  <%
					  rs_T.movenext
					  Next
					  %>
                      </select>
					  </td>
                    </tr>
                    <tr>
                      <td height="34" align="center" style="padding-left:10px">主题：</td>
                      <td class="word_grey"><input name="subject" type="text" id="subject" size="50"></td>
                    </tr>
                    <tr>
                      <td width="16%" height="90" align="center">表情：</td>
                      <td width="84%"><INPUT name=emote type=radio class="noborder" value=0 checked>
                        <IMG height=20 
            src="Images/Face/face0.gif" width=20 align=middle border=0>
                        <INPUT name=emote type=radio class="noborder" value=1>
                        <IMG height=19 
            src="Images/Face/face1.gif" width=19 align=middle border=0>
                        <INPUT name=emote type=radio class="noborder" value=2>
                        <IMG height=20 
            src="Images/Face/face2.gif" width=20 align=middle border=0>
                        <INPUT name=emote type=radio class="noborder" value=3>
                        <IMG height=20 
            src="Images/Face/face3.gif" width=20 align=middle border=0>
                        <INPUT name=emote type=radio class="noborder" value=4>
                        <IMG height=20 
            src="Images/Face/face4.gif" width=20 align=middle border=0>
                        <INPUT name=emote type=radio class="noborder" value=5>
                        <IMG height=20 
            src="Images/Face/face5.gif" width=20 align=middle border=0>
                        <INPUT name=emote type=radio class="noborder" value=6>
                        <IMG height=20 
            src="Images/Face/face6.gif" width=20 align=middle border=0>
                        <INPUT name=emote type=radio class="noborder" value=7>
                        <img height=20 
            src="Images/Face/face7.gif" width=20 align=middle border=0><br>
                        <INPUT name=emote type=radio class="noborder" value=8>
                        <IMG height=20 
            src="Images/Face/face8.gif" width=20 align=middle border=0>
                        <INPUT name=emote type=radio class="noborder" value=9>
                        <IMG height=20 
            src="Images/Face/face9.gif" width=20 align=middle border=0> 
			            <INPUT name=emote type=radio class="noborder" value=10>
                        <IMG height=20 
            src="Images/Face/face10.gif" width=20 align=middle border=0>
                        <INPUT name=emote type=radio class="noborder" value=11>
                        <IMG height=20 
            src="Images/Face/face11.gif" width=20 align=middle border=0>
                        <INPUT name=emote type=radio class="noborder" value=12>
                        <IMG height=20 
            src="Images/Face/face12.gif" width=20 align=middle border=0>
                        <INPUT name=emote type=radio class="noborder" value=13>
                        <IMG height=20 
            src="Images/Face/face13.gif" width=20 align=middle border=0>
                        <INPUT name=emote type=radio class="noborder" value=14>
                        <IMG height=20 
            src="Images/Face/face14.gif" width=20 align=middle border=0>
                        <INPUT name=emote type=radio class="noborder" value=15>
                        <img height=20 
            src="Images/Face/face15.gif" width=20 align=middle border=0><br>
                        <INPUT name=emote type=radio class="noborder" value=16>
                        <IMG height=20 
            src="Images/Face/face16.gif" width=20 align=middle border=0>
                        <INPUT name=emote type=radio class="noborder" value=17>
                        <IMG height=20 
            src="Images/Face/face17.gif" width=20 align=middle border=0>
                        <INPUT name=emote type=radio class="noborder" value=18>
                        <IMG height=20 
            src="Images/Face/face18.gif" width=20 align=middle border=0>
                        <INPUT name=emote type=radio class="noborder" value=19>
                        <IMG height=19 
            src="Images/Face/face19.gif" width=19 align=middle border=0>
                        <INPUT name=emote type=radio class="noborder" value=20>
                        <IMG height=19 
            src="Images/Face/face20.gif" width=19 align=middle border=0>
                        <INPUT name=emote type=radio class="noborder" value=21>
                        <IMG height=19 
            src="Images/Face/face21.gif" width=19 align=middle border=0>
                        <INPUT name=emote type=radio class="noborder" value=22>
                        <IMG height=19 
            src="Images/Face/face22.gif" width=19 align=middle border=0>
                        <INPUT name=emote type=radio class="noborder" value=23>
                        <IMG height=19 
            src="Images/Face/face23.gif" width=19 align=middle border=0></td>
                    </tr>
                    <tr>
                      <td height="46" colspan="2" align="center" style="padding-left:10px"><table width="70%" height="25"  border="0" cellpadding="-2" cellspacing="-2">
						<tr>
                          <td width="21%"><img src="Images/UBB/B.gif" width="21" height="20" onClick="bold()">&nbsp;<img src="Images/UBB/I.gif" width="21" height="20" onClick="xt()">&nbsp;<img src="Images/UBB/U.gif" width="21" height="20" onClick="ul()"></td>
                          <td width="79%">字体
                            <select name="font" class="wenbenkuang" id="font" onChange="sf(this.options[this.selectedIndex].value)">
                              <option value="宋体" selected>宋体</option>
                              <option value="黑体">黑体</option>
                              <option value="隶书">隶书</option>
                              <option value="楷体">楷体</option>
                            </select>
                            字号<span class="pt9">
                            <SELECT 
      name=size class="wenbenkuang" onChange="ss(this.options[this.selectedIndex].value)">
                              <OPTION value=1>1</OPTION>
                              <OPTION value=2>2</OPTION>
                              <OPTION 
        value=3 selected>3</OPTION>
                              <OPTION value=4>4</OPTION>
                              <option value="5">5</option>
                              <option value="6">6</option>
                              <option value="7">7</option>
                            </SELECT>
颜色 <select onChange="sc(this.options[this.selectedIndex].value)" name="color" size="1" class="wenbenkuang" id="select">
            <option selected>默认颜色</option>
            <option style="color:#FF0000" value="#FF0000">红色热情</option>
            <option style="color:#0000FF" value="#0000ff">蓝色开朗</option>
            <option style="color:#ff00ff" value="#ff00ff">桃色浪漫</option>
            <option style="color:#009900" value="#009900">绿色青春</option>
            <option style="color:#009999" value="#009999">青色清爽</option>
            <option style="color:#990099" value="#990099">紫色拘谨</option>
            <option style="color:#990000" value="#990000">暗夜兴奋</option>
            <option style="color:#000099" value="#000099">深蓝忧郁</option>
            <option style="color:#999900" value="#999900">卡其制服</option>
            <option style="color:#ff9900" value="#ff9900">镏金岁月</option>
            <option style="color:#0099ff" value="#0099ff">湖波荡漾</option>
            <option style="color:#9900ff" value="#9900ff">发亮蓝紫</option>
            <option style="color:#ff0099" value="#ff0099">爱的暗示</option>
            <option style="color:#006600" value="#006600">墨绿深沉</option>
            <option style="color:#999999" value="#999999">烟雨蒙蒙</option>
        </select>                           </span></td>
                        </tr>
                      </table></td>
                      </tr>
                    <tr>
                      <td height="145" align="center">内容：</td>
                      <td><textarea name="content" cols="60" rows="14" class="wenbenkuang" id="content"
					   onkeydown=CountStrByte(this.form.content,this.form.total,this.form.used,this.form.remain);
					    onkeyup=CountStrByte(this.form.content,this.form.total,this.form.used,this.form.remain);></textarea></td>
                    </tr>
                    <tr>
                      <td height="33" align="center" style="padding-left:10px"><p>字节：</p>                        </td>
                      <td style="padding-left:10px" class="word_grey">最多允许 
                        <input name="total" type="text" disabled class="noborder" id="total"  value="1600" size="4"> 
                        个字节 已用字节：&nbsp;
                        <input name="used" type="text" disabled class="noborder" id="used"  value="0" size="4">                        
                        剩余字节：
                        <input name="remain" type="text" disabled class="noborder" id="remain" value="1600" size="4">
                        <input name="HIP" type="hidden" id="HIP" value="<%=Request.ServerVariables("REMOTE_ADDR")%>"></td>
                    </tr>
                    <tr>
                      <td height="34" align="center" style="padding-left:10px">&nbsp;</td>
                      <td class="word_grey"><input name="Button" type="button" class="btn_grey" value="提交主题信息" onClick="check()">                          
                        <input name="Submit2" type="reset" class="btn_grey" value="重写主题信息"></td></tr>
                  </table>
				    </form>
				  </td>
                </tr>
				<tr>
				<td height="5"></td>
				</tr>
              </table>			  </td>
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
<%
rs.close()
Set rs=Nothing
%>
