<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/conn.asp" -->
<!--#include file="../../NetCustomer/F_Display_z.asp" -->
<%
'查询全部商品信息
Set rs_sp = Server.CreateObject("ADODB.Recordset")
sql_SP = "SELECT spnumber, spname,pack FROM DB_SPInfo"
rs_sp.open sql_SP,conn,1,3
'查询业务员信息
Set rs_user = Server.CreateObject("ADODB.Recordset")
sql_W = "SELECT ywynumber, name, workqy FROM DB_WorkerInfo WHERE name = '"&session("username")&"'"
rs_user.open sql_W,conn,1,3
%>
<%'保存销售信息
if request.Form("spNO")<>"" then
spNO=request.Form("spNO")
khNO=request.Form("khNO")
pack=request.Form("pack")
Ddate=request.Form("date")
sellsum=request.Form("sellsum")
price=request.Form("price")
zmoney=request.Form("zmoney")
userno=request.Form("userno")
Ins="Insert into DB_Sell (spNO,khNO,pack,D_date,sellsum,price,zmoney,userno) values('"&_
spNO&"','"&khNO&"','"&pack&"','"&Ddate&"',"&sellsum&","&price&","&zmoney&",'"&userno&"')"
conn.execute(Ins)
 %>
	<script language="javascript">
	  alert("销售登记成功！")
	  window.location.href="default.asp";
	</script>
<%end if%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>销售信息查询！</title>
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
<script language="javascript">
function mycheck(){
if (form1.khNO.value=="0"){alert("您暂时还没有客户，请与管理员联系！");return}
if (form1.sellsum.value==0){alert("请输入销售数量！");form1.sellsum.focus();return}
form1.submit();
}
</script>
<body>
<table width="595" height="242" border="0" cellpadding="-2" cellspacing="-2"
 background="../Images/manage.gif">
  <tr>
    <td valign="top"><table width="595" height="400" cellpadding="-2" cellspacing="-2">
      <tr>
        <td width="10" height="94">&nbsp;</td>
        <td width="577">&nbsp;</td>
        <td width="10"><span class="style2"></span></td>
      </tr>
      <tr>
        <td height="258">&nbsp;</td>
        <td valign="top">          <div align="center">
            <table width="97%" height="31" cellpadding="-2"  cellspacing="-2">
              <tr>
                <td width="4%"><div align="right"><img src="../Images/add.gif" width="18" height="18"></div></td>
                <td width="96%"><div align="left">&nbsp;业务员：<%=session("UserName")%>
				 　　　工作区域：<%qyNO=rs_user("workqy")
				                   call display(qyNO)%> </div></td>
              </tr>
            </table>
             <form ACTION="sell_add.asp" METHOD="POST" name="form1">
              <table width="456" align="center">
                <tr valign="baseline">
                  <td width="88" nowrap><div align="center">商品名称：</div></td>
                  <td colspan="3" valign="baseline">
<%'显示所选商品的基本信息
SelectID=request.QueryString("SPID")
if SelectID<>"" then
	sql="select* from DB_spinfo where spnumber='"&SelectID&"'"
	set rs_sp_s=conn.execute(sql)
else
	sql="select* from DB_spinfo where spnumber='"&rs_sp("spnumber")&"'"
	set rs_sp_s=conn.execute(sql)
end if%>
                <div align="left">
<script language="javascript">
function ChangeItem(){//当用户改变商品名称时调用该函数
var SPID=form1.spNO.value;
window.location.href="sell_add.asp?SPID="+SPID;
}
</script>
<select name="spNO" id="spNO" onChange="ChangeItem()">
<%While (NOT rs_sp.EOF)%>
    <option value="<%=(rs_sp("spnumber"))%>"
	<% if rs_sp("spnumber")=SelectID then %>selected<% end if%>>
		<%=(rs_sp("spname"))%> [<%=(rs_sp("pack"))%>]</option>
        <% rs_sp.MoveNext()
Wend%>
</select>                    
                 </div></td>
                </tr>
                <tr valign="baseline">
                  <td nowrap><div align="center">客户名称：</div></td>
<% '创建查询该业务员的所有客户的记录集
Set rs_kh = Server.CreateObject("ADODB.Recordset")
ReDim areaID(255)
sql_kh_1="select* from Q_area where father="& rs_user("workqy")
set rs_kh_1=conn.execute(sql_kh_1)
areaID(0)=rs_user("workqy")
i=1
while not rs_kh_1.eof
	areaID(i)=rs_kh_1("ID")
	sql_kh_1_1="select* from Q_area where father="& rs_kh_1("ID")
	set rs_kh_1_1=conn.execute(sql_kh_1_1)
	while not rs_kh_1_1.eof
		sql_kh_1_1_1="select* from Q_area where father="& areaID(i)
		set rs_kh_1_1_1=conn.execute(sql_kh_1_1_1)
	
		while not rs_kh_1_1_1.eof
			i=i+1
			areaID(i)=rs_kh_1_1_1("ID")
			rs_kh_1_1_1.movenext
		wend
		rs_kh_1_1.movenext
	wend
	rs_kh_1.movenext
	i=i+1
wend
for i=0 to UBound(areaID)
if areaID(i)<>"" then
	str=str+"qyname="& areaID(i) &" or "
	strTJ=left(str,len(str)-4)
end if
next
sql_kh = "SELECT khnumber, khname, qyname FROM DB_KhInfo where "& strTJ 
rs_kh.open sql_kh,conn,1,3
%>
				  <td colspan="3" valign="baseline">
                    <div align="left">
					<% if not rs_kh.eof then
					rs_kh.movefirst %>
					   <select name="khNO" id="khNO">
					   <% while not rs_kh.eof %>
                        <option value="<%=rs_kh("khnumber")%>"><%=rs_kh("khname")%></option>
					  <% rs_kh.movenext
						wend %>
						</select>
					  <% else %>
					  <select name="khNO" id="khNO">
                        <option value="0">----</option>
                      </select>
					  <%end if%>                   
                        </div></td>
                </tr>
                <tr valign="baseline">
                  <td nowrap><div align="center">包　　装：</div></td>
                  <td colspan="3" valign="baseline">
                    <div align="left">
                      <input name="pack" type="text" class="Sytle_auto" value="<%=rs_sp_s("pack")%>" size="42" readonly="yes">
                    </div></td>
                </tr>
                <tr valign="baseline">
                  <td nowrap><div align="center">销售日期：</div></td>
                  <td valign="baseline">
                    <div align="left">
                      <input name="date" type="text" class="Sytle_text" readonly="yes" value="<%=date()%>" size="32">
                    </div></td>
                  <td valign="baseline">销售数量：</td>
                  <td valign="baseline"><div align="left">
				  <script language="javascript">
				  function auto(){//自动计算金额
				  form1.zmoney.value=form1.price.value*form1.sellsum.value;
				  }
				  </script>
                    <input name="sellsum" type="text" class="Sytle_text" value="0" onChange="auto()">
                  </div></td>
                </tr>
                <tr valign="baseline">
                  <td nowrap><div align="center">单　　价：</div></td>
                  <td width="99" valign="baseline">
                    <div align="left">
                      <input name="price" type="text" class="Sytle_text" value="<%=rs_sp_s("price")%>" size="32" readonly="yes">
                    </div></td>
                  <td width="69" valign="baseline">金　　额：</td>
                  <td width="180" valign="baseline"><div align="left">
                    <input name="zmoney" type="text" class="Sytle_text" value="0" size="32" readonly="yes">
                  </div></td>
                </tr>
                <tr valign="middle">
                  <td height="49" colspan="4" nowrap><div align="center">
                      <input name="Button" type="button" class="Style_button" value="保存" onClick="mycheck()">
                      <input name="Reset" type="reset" class="Style_button" value="重置">
                      <input name="Button" type="button" class="Style_button" value="返回" onClick="JScript:history.back();">
                      <input name="userno" type="hidden" id="userno" value="<%=rs_user("ywynumber")%>">
                  </div></td>
                  </tr>
              </table>
             </form>
            <p>&nbsp;</p>
          </div></td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
</html>
