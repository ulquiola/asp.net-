<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn/conn.asp"-->
<%
function filter_Str(InString)
'***********************************
'功能：过滤输入字符串中的危险符号
'调用方法：filter_Str("String")
'***********************************
	NewStr=Replace(InString,"'","''")
	NewStr=Replace(NewStr,"<","&lt;")
	NewStr=Replace(NewStr,">","&gt;")
	NewStr=Replace(NewStr,"chr(60)","&lt;")
	NewStr=Replace(NewStr,"chr(37)","&gt;")
	NewStr=Replace(NewStr,"""","&quot;")
	NewStr=Replace(NewStr,";",";;")
	NewStr=Replace(NewStr,"--","-")
	NewStr=Replace(NewStr,"/*"," ")
	NewStr=Replace(NewStr,"%"," ")
	filter_Str=NewStr
end function
session.Timeout=120 
if Request.Form("usename")<>"" and request.Form("pwd")<>"" then
	session("usename")=filter_Str(request.Form("usename"))
	session("pwd")=filter_Str(request.Form("pwd"))
	session("purview")=filter_Str(request.Form("purview"))
	set rs=server.CreateObject("adodb.recordset")
	sql="select * from tb_user where usename='"&session("usename")&"'"
    rs.open sql,conn,1,3
	set rs=conn.execute(sql)
		if rs.eof and rs.bof then
		%>
			<script language="javascript">
			alert("您输入的用户名称错误，请重新输入！");
			location='index.asp';
			 </script> 
				<%
				elseif rs("pwd")<>session("pwd") then
				%>
				<script language="javascript">
				alert("您输入的用户密码错误，请重新输入！");
				location='index.asp';
				</script>
				<%else 
					 if rs("purview")=session("purview") then
				 %>
				 <script language="javascript">
					alert("欢迎您登录本网站！");
					window.location.href='main.asp';
					</script>
					 <%else%>
						<script language="javascript">
						alert("您的用户级别不正确，请重新选择！");
						location='index.asp';
						</script>
			<%
			end if
		end if
	end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>用户登录</title>
<link rel="stylesheet" href="Css/style.css">
<style type="text/css">
<!--
body {
	margin-top: 0px;
	margin-bottom: 0px;
	background-image: url(images/bg.gif);
}
.STYLE1 {color: #000000}
.STYLE8 {color: #FF9900}
-->
</style></head>
<script language="javascript">
function Mycheck(form1)
{
  for(i=0;i<form1.length;i++)
   {
    if(form1.elements[i].value=="")
	{alert(form1.elements[i].title + "不能为空!");return false;}
    }
       form1.submit();
}
</script>
<body>
<p>&nbsp;</p>
<p>&nbsp;</p>
<table width="775" height="453" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="601" height="453" background="images/zhece.gif"><form action="" method="post" name="form1">
      <table width="740" height="241" border="0" align="right" cellpadding="0" cellspacing="0">
      <tr>
        <td width="259" height="169" valign="top">&nbsp;</td>
        <td width="481" valign="top"><table width="265" height="124" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td width="89"><div align="center" class="STYLE1">
                <div align="right">用户名：　 </div>
            </div></td>
            <td width="201"><input name="usename" type="text" id="usename" onKeyDown="if(event.keyCode==13){form1.password.focus();}" title="用户名"></td>
          </tr>
          <tr>
            <td><div align="center" class="STYLE1">
                <div align="right" class="STYLE1">密　码：　</div>
            </div></td>
            <td><input name="pwd" type="password" id="pwd" onKeyDown="if(event.keyCode==13){form1.yan.focus();}" title="密码"></td>
          </tr>
          <tr>
            <td><div align="center" class="STYLE1">
                <div>&nbsp; 用户级别：</div>
            </div></td>
            <td><select name="purview" class="STYLE1" id="purview">
                <option value="普通用户">普通用户</option>
                <option value="系统用户">系统用户</option>
            </select></td>
          </tr>
          <tr>
            <td colspan="2"><div align="center"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                    <input name="Submit" type="button" class="btn_grey" onClick="Mycheck(form1)" value="登录">
              
              &nbsp;
              <input name="Submit2" type="reset" class="btn_grey" value="重置">
              <a href="user_zc.asp">&nbsp;<span class="STYLE8">[用户注册]</span></a></div></td>
          </tr>
        </table></td>
      </tr>
    </table>
	</form>
	</td>
  </tr>
</table>
</body>
</html>
<%id =session("usename")%>
