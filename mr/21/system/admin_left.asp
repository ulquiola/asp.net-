<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn/conn.asp"-->
<%
set rs=server.CreateObject("adodb.recordset")
sql="select * from tb_user"
rs.open sql,conn,1,3
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>LEFT</title>
<link href="Img/Menu/Style.css" rel="stylesheet" type="text/css">
<style type="text/css">
a:link                  { text-decoration: none; color: #3A4273 }
a:hover                 { text-decoration: underline }
a:visited               { text-decoration: none; color: #464F86 }
body, table, tr, td	{ scrollbar-base-color: #E3E3EA; scrollbar-arrow-color: #007db5; font-size: 12px }
.header			{ color: #F1F3FB; background-color: #CCCCCC; font-family: Tahoma,Verdana; font-weight: bold; font-size: 12px }
.tablerow		{ font-family: Tahoma,Verdana; color: #464F86; font-size: 12px }
BODY {
scrollbar-face-color:#BBCEE8;
scrollbar-highlight-color:#F2FAFD;
scrollbar-3dlight-color:#BCCFE9;
scrollbar-darkshadow-color:#BBCEF8;
scrollbar-shadow-color:#85A7D6;
scrollbar-arrow-color:#2445C8;
scrollbar-track-color:#BCD9F3;
}
</style>
<script language="JavaScript">
function chgTDColor(oTD) {
	oTD.style.cursor = "hand";
	if(oTD.style.backgroundColor == "") {
		oTD.style.backgroundColor = "#E3E3EA";
		if(oTD.className == "header") {
			oTD.style.color = "#464F86";
		}
	} else {
		oTD.style.backgroundColor = "";
		if(oTD.className == "header") {
			oTD.style.color = "#F1F3FB";
		}
	}
}

function showhideMenu(n) {
	var oTR,oImg;
	
	oTR = eval("document.all.tr"+n);
	oImg = eval("img"+n);
	if (oTR.style.display == "none") {
		oTR.style.display = "block";
		oImg.src="images/Menu2.gif";
		oImg.parentElement.parentElement.bgColor="#E3E3EA"
		for (i=1;i<=3;i++ ) {
			if (i!=n) {
				var mTR = eval("tr"+i);
				var mImg = eval("img"+i);
				mImg.parentElement.parentElement.bgColor="";
			}
		}
	} else {
		oTR.style.display = "none";
		oImg.src="images/Menu2.gif";
	}
}
</script>
</head>
<base target="EwSys_Main" oncontextmenu=self.event.returnValue=false>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0" bgcolor="#EEEEF6">
  <tr>
    <td valign="top">
	  <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#EEEEF6" class="tablerow">
        <tr> 
          <td background="images/Menu3.gif" class="header" onClick="showhideMenu(3)" onMouseOver="chgTDColor(this)" onMouseOut="chgTDColor(this)"><img src="images/Menu2.gif" align="absmiddle" width="25" height="25" name="img3"><b> 
            服务器情况监测 </b></td>
        </tr>
        <tr> 
          <td height="1" bgcolor="#000000"></td>
        </tr>
        <tr style="display: none" id="tr3"> 
          <td> 
		  	<table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td height="22" bgcolor="#EEEEF6" onClick="top.EwSys_Main.location.href='server.asp';" onMouseOver="chgTDColor(this)" onMouseOut="chgTDColor(this)">&nbsp;&nbsp;&nbsp;&nbsp;<img src="images/Menu4.gif" border="0" align="absmiddle"> 服务器信息</td>
              </tr>
              <tr> 
                <td height="22" bgcolor="#EEEEF6" onClick="top.EwSys_Main.location.href='subassembly.asp';" onMouseOver="chgTDColor(this)" onMouseOut="chgTDColor(this)">&nbsp;&nbsp;&nbsp;&nbsp;<img src="images/Menu4.gif" border="0" align="absmiddle"> 服务器组件信息</td>
              </tr>
            </table>
		  </td>
        </tr>
        <tr> 
          <td background="images/Menu3.gif" class="header" onClick="showhideMenu(5)" onMouseOver="chgTDColor(this)" onMouseOut="chgTDColor(this)"><img src="images/Menu2.gif" align="absmiddle" width="25" height="25" name="img5"><b><strong>服务器磁盘监测</strong></b></td>
        </tr>
        <tr> 
          <td height="1" bgcolor="#000000"></td>
        </tr>
        <tr id="tr5" style="display: none"> 
          <td> 
		  	<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
              <tr bgcolor="#003399" class="tablerow"> 
                <td height="22" bgcolor="#EEEEF6" onClick="top.EwSys_Main.location.href='drive.asp';" onMouseOver="chgTDColor(this)" onMouseOut="chgTDColor(this)">&nbsp;&nbsp;&nbsp;&nbsp;<img src="images/Menu4.gif" border="0" align="absmiddle"> 磁盘信息</td>
              </tr>
            </table>
		  </td>
        </tr>
		<tr> 
          <td background="images/Menu3.gif" class="header" onClick="showhideMenu(2)" onMouseOver="chgTDColor(this)" onMouseOut="chgTDColor(this)"><img src="images/Menu2.gif" align="absmiddle" width="25" height="25" name="img2"><b><strong>磁盘文件操作速度监测</strong></b></td>
        </tr>
		<tr> 
          <td height="1" bgcolor="#000000"></td>
        </tr>
        <tr id="tr2" style="display: none"> 
          <td height=""> 
		  	<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" >
			  <tr bgcolor="#003399"> 
                <td height="22" bgcolor="#EEEEF6" onClick="top.EwSys_Main.location.href='speed.asp';" onMouseOver="chgTDColor(this)" onMouseOut="chgTDColor(this)">&nbsp;&nbsp;&nbsp;&nbsp;<img src="images/Menu4.gif" border="0" align="absmiddle"> 磁盘文件操作速度</td>
              </tr> 
            </table>
		  </td>
        </tr>
        <tr> 
          <td background="images/Menu3.gif" class="header" onClick="showhideMenu(1)" onMouseOver="chgTDColor(this)" onMouseOut="chgTDColor(this)"><img src="images/Menu2.gif" align="absmiddle" width="25" height="25" name="img1"><b><strong><b><b><strong>会 员 管 理</strong></b></b></strong></b></td>
        </tr>
        <tr> 
          <td height="1" bgcolor="#000000"></td>
        </tr>
        <tr id="tr1"> 
          <td> 
		  	<table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr bgcolor="#003399" class="tablerow"> 
				  <%if session("purview")<>"系统用户" then%>
                <td height="22" bgcolor="#EEEEF6" onClick="JScript:alert('对不起，您没有该级别!');" onMouseOver="chgTDColor(this)" onMouseOut="chgTDColor(this)"> &nbsp;&nbsp;&nbsp;&nbsp;<img src="images/Menu4.gif" border="0" align="absmiddle"> 会员管理</td>
				  <%else%>
                <td height="22" bgcolor="#EEEEF6" onClick="top.EwSys_Main.location.href='adm_Manage.asp';" onMouseOver="chgTDColor(this)" onMouseOut="chgTDColor(this)"> &nbsp;&nbsp;&nbsp;&nbsp;<img src="images/Menu4.gif" border="0" align="absmiddle"> 会员管理</td>
		  <%end if%>
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
