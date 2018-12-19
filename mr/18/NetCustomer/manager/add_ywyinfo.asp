<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Connections/conn.asp" -->
<!--#include file="../../NetCustomer/F_Display.asp" -->
<%
session("userno")=session("autoNO")
if request.Form("name")<>"" then
'判断输入的业务员是否存在
  Set rs_W = Server.CreateObject("ADODB.Recordset")
  sql_W="SELECT name FROM DB_WorkerInfo WHERE name='"&request.Form("name")&"'"
  rs_W.open sql_W,conn,1,3
  if rs_W.eof then  '如果不存在则插入业务员信息
  cname=Request.Form("name")
  sex=Request.Form("sex")
  folk=Request.Form("folk")
  birthday=Request.Form("birthday")
  rcdate=Request.Form("rcdate")
  xl=Request.Form("xl")
  sfzNO=Request.Form("sfzNO")
  zy=Request.Form("zy")
  tel=Request.Form("tel")
  workqy=Request.Form("workqy")
  address=Request.Form("address")
  email=Request.Form("email")
  memo=Request.Form("memo")
  Ins="Insert Into DB_WorkerInfo (ywynumber,name,sex,folk,birthday,rcdate,xl,sfzNO,"&_
  "zy,tel,workqy,address,email,memo) values('"&session("autoNO")&"','"&cname&"','"&sex&_
  "','"&folk&"','"&birthday&"','"&rcdate&"','"&xl&"','"&sfzNO&"','"&zy&"','"&tel&"','"&_
  workqy&"','"&address&"','"&email&"','"&memo&"')"
  conn.execute(Ins)
  Ins_P="Insert Into DB_Purview (UserNO) values('"&session("autoNO")&"')"
  conn.execute(Ins_P)
%>
 	<script language="javascript">
	alert("业务员信息添加成功！");
	window.close();
	</script>
<% response.end
else%>
 	<script language="javascript">
	alert("该业务员信息已经存在！");
	window.close();
	</script>
<% end if
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>业务员信息！</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style.css" rel="stylesheet">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	background-image: url(../Images/bg.gif);
}
.style1 {color: #FFFFFF}
.style2 {color: #a2bcc5}
-->
</style></head>

<body>
<script language="javascript">
function mycheck(){
if (form1.name.value=="")
{alert("请输入业务员姓名！");form1.name.focus();return;}
if(form1.sfzNO.value=="")
{alert("请输入身份证号码！");form1.sfzNO.focus();return;}
form1.submit();
}
</script>
<% 
'查询业务员信息数据表中最大的业务员编号
Set rs_Max = Server.CreateObject("ADODB.Recordset")
sql_max = "SELECT Max(ywynumber) as MaxNO  FROM DB_Workerinfo"
rs_Max.open sql_max,conn,1,3
Dim no,valno,mymonth,cmonth,cday
no=trim(rs_Max("MaxNO"))
'生成五位的编号
select case len(int(Right(no,5)+1))
	case 1
		cno="0000"+Cstr(int(Right(no,5)+1))
	case 2
		cno="000"+Cstr(int(Right(no,5)+1))
	case 3
		cno="00"+Cstr(int(Right(no,5)+1))
	case 4
		cno="0"+Cstr(int(Right(no,5)+1))
	case 5
		cno=Cstr(int(Right(no,5)+1))
	case Else
		cno="00001"
end select
intno="YWY"+cno '组成新的数据编号
session("autoNO")=intno  '将生成的数据编号赋值给会话变量
%>
<table width="595" height="242" border="0" cellpadding="-2" cellspacing="-2"
 background="../Images/manage.gif">
  <tr>
    <td valign="top"><table width="595" cellpadding="-2" cellspacing="-2">
      <tr>
        <td width="10" height="112">&nbsp;</td>
        <td width="577">&nbsp;</td>
        <td width="10"><span class="style2"></span></td>
      </tr>
      <tr>
        <td height="258">&nbsp;</td>
        <td valign="top"><div align="center">
          <form ACTION="add_ywyinfo.asp" METHOD="POST" name="form1">
            <table border="1" align="center" cellpadding="-2" cellspacing="-2"
			 bordercolor="#669999" bordercolordark="#FFFFFF">
              <tr valign="middle">
                <td width="72" align="right" nowrap bgcolor="#809EA4">
				<div align="center" class="style1">编　　号：</div></td>
                <td width="105">                    <div align="left">
                  <input name="ywynumber" type="text" class="Sytle_text" id="ywynumber"
				   value="<%=session("autoNO")%>" size="32" readonly="yes">
                </div></td>
                <td width="71" bgcolor="#809EA4"><div align="center" class="style1">姓　　名：</div></td>
                <td width="96"><div align="left">
                  <input name="name" type="text" class="Sytle_text" value="" size="32">
                </div></td>
                <td width="68" bgcolor="#809EA4"><div align="center" class="style1">性　　别：</div></td>
                <td width="97"><div align="left">
                  <select name="sex" id="select3">
                      <option value="男" selected>男</option>
                      <option value="女">女</option>
                  </select>
                </div></td>
              </tr>
              <tr valign="middle">
                <td align="right" nowrap bgcolor="#809EA4"><div align="center" class="style1">民　　族：</div></td>
                <td><div align="left">
                  <input name="folk" type="text" class="Sytle_text" value="汉族" size="32">
                </div></td>
                <td bgcolor="#809EA4"><div align="center" class="style1">生　　日：</div></td>
                <td><div align="left">
                  <input name="birthday" type="text" class="Sytle_text" value="<%=date()-20*365%>" size="32">
                </div></td>
                <td bgcolor="#809EA4"><div align="center" class="style1">入厂日期：</div></td>
                <td><div align="left">
                  <input name="rcdate" type="text" class="Sytle_text" value="<%=date()%>" size="32">
                </div></td>
              </tr>
              <tr valign="middle">
                <td align="right" nowrap bgcolor="#809EA4"><div align="center" class="style1">学　　历：</div></td>
                <td><div align="left">
                  <select name="xl" id="select4">
                      <option value="博士及以上" selected>博士及以上</option>
                      <option value="硕士">硕士</option>
                      <option value="研究生">研究生</option>
                      <option value="本科">本科</option>
                      <option value="专科">专科</option>
                      <option value="高中及中专">高中及中专</option>
                      <option value="初中及以下">初中及以下</option>
                  </select>
                </div></td>
                <td bgcolor="#809EA4"><div align="center" class="style1">身份证号：</div></td>
                <td colspan="3"><div align="left">
                  <input name="sfzNO" type="text" class="Sytle_auto" size="32">
                </div></td>
              </tr>
              <tr valign="middle">
                <td align="right" nowrap bgcolor="#809EA4"><div align="center" class="style1">专　　业：</div></td>
                <td><div align="left">
                  <select name="zy" id="select5">
                      <option value="计算机" selected>计算机</option>
                      <option value="食品">食品</option>
                      <option value="机械制造">机械制造</option>
                      <option value="医生">医生</option>
                      <option value="护士">护士</option>
                      <option value="教师">教师</option>
                      <option value="其他">其他</option>
                  </select>
                </div></td>
                <td bgcolor="#809EA4"><div align="center" class="style1">联系电话：</div></td>
                <td colspan="3"><div align="left">
                  <input name="tel" type="text" class="Sytle_auto" value="" size="32">
                </div></td>
                </tr>
              <tr valign="middle">
                <td align="right" nowrap bgcolor="#809EA4"><div align="center" class="style1">工作区域：</div></td>
                <td colspan="5"><div align="left">                  <table width="98%"  cellspacing="-2" cellpadding="-2">
                    <tr>
                      <td width="86%">
						<% qyNO=request("qyNO")
						call Display(qyNO)  '显示工作区域%>
					  </td>
                      <td width="14%"><input name="workqy" type="hidden" id="workqy4" value="<%=request("qyNO")%>"></td>
                    </tr>
                  </table>
				</div></td>
                </tr>
              <tr valign="middle">
                <td align="right" nowrap bgcolor="#809EA4"><div align="center" class="style1">地　　址：</div></td>
                <td colspan="5">
                  <div align="left">
                    <input name="address" type="text" class="Style_upload" value="" size="32">                  
                  </div></td>
              </tr>
              <tr valign="middle">
                <td align="right" nowrap bgcolor="#809EA4"><div align="center" class="style1">Email:</div></td>
                <td colspan="5">
                  <div align="left">
                    <input name="email" type="text" class="Sytle_auto" value="" size="32">                
                  </div></td>
              </tr>
              <tr valign="middle">
                <td align="right" nowrap bgcolor="#809EA4"><div align="center" class="style1">备　　注：</div></td>
                <td colspan="5">
                  <div align="left">
                    <input name="memo" type="text" class="Style_upload" value="" size="32">                
                  </div></td>
              </tr>
              <tr valign="middle">
                <td height="35" colspan="6" nowrap><div align="center">
                    <input name="Button" type="button" class="Style_button" value="保存"
					 onClick="mycheck()">                
                    <input name="Reset" type="reset" class="Style_button" value="重置">
                    <input name="Submit2" type="button" class="Style_button" value="关闭"
					 onClick="JScript:window.close()">
                </div></td>
                </tr>
            </table>
          </form>
          <p>&nbsp;</p>
        </div></td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="21">&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
</html>

