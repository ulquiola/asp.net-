<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/conn.asp"-->
<%
set rsnei=server.createobject("adodb.recordset")
rs2="select * from tb_project where type1='国内'"
rsnei.open rs2,conn,1,3
set rswai=server.CreateObject("adodb.recordset")
rs3="select * from tb_project where type1='国外'"
rswai.open rs3,conn,1,3
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>明日在线搜索</title>
<style type="text/css">
<!--
.style1 {font-size: 9px}
.style2 {font-size: 9pt}
.style3 {font-size: 10pt}
.style4 {color: #FF0000}
.STYLE6 {font-size: 9pt; color: #000000; }
-->
</style>
</head>
<body onLoad="clockon()" background="images/bg.gif"> 
<!--#include file="shangbiao.asp"-->
<table width="780" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#FFFFFF" bgcolor="#FFFFFF" algin="center"> 
     <td width="23" height="607" align="right"   valign="top"></td>
     <td width="756" valign="top"> 
    <table width="780" height="557" border="0" cellpadding="0" cellspacing="0" >
              <td width="780" height="553" align="center" valign="top"><table width="780" height="51" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="21" background="images/sousou2.gif">&nbsp;</td>
          <td width="81" height="47" background="images/tubiao.gif">&nbsp;</td>
          <td width="7" height="47" valign="bottom" background="images/sousou2.gif" bgcolor="#FFFFFF">&nbsp;</td>
          <td width="103" valign="bottom" background="images/ITqiyegushi.gif" bgcolor="#FFFFFF">&nbsp;</td>
          <td width="390" align="right" valign="bottom" background="images/sou2.gif" bgcolor="#FFFFFF"><a href="index.asp"style="color:#3661a6"></a></td>
          <td width="78" align="right" valign="bottom" background="images/sou2.gif" bgcolor="#FFFFFF"><div align="center"><a href="index.asp" class="style2"style="color:#3661a6">返回首页</a></div></td>
          <td width="100" valign="bottom" background="images/sousouhou.gif" bgcolor="#FFFFFF"><a href="index.asp"style="color:#3661a6"></a></td>
        </tr>
        <tr>
          <td height="3" colspan="7" background="images/erjiye1.gif"></td>
        </tr>
        <tr>
          <td height="30" colspan="7">&nbsp;</td>
        </tr>
      </table>
         <table width="675" height="97" border="0" align="center" cellpadding="0" cellspacing="0"style="color:#336699" >
    <tr>
      <td height="21"><div align="center" class="style3"><strong>国内IT企业故事</strong></div></td>
            </tr>
    <tr>
      <td height="14">
        <span class="style4"></div>
          </span>              <div align="center"></div></td>
              </tr>
    <% '分页'
	  rsnei.pagesize=6
	  page=clng(request("page"))
	  if page<1 then page=1
	  rsnei.absolutepage=page
	  for i=1 to rsnei.pagesize
	  jj=rsnei("introduce")
	  %>
    <tr>
      <td height="45"><span class="STYLE6"><a href="#"onClick="javascript:window.open('introduce7.asp?id=<%=rsnei("ID")%>')"><%=rsnei("name1")%></a><br>
            <% if len(jj)>200 then%>
            <%=replace(left(jj,200)&"....",chr(13),"<br>")%>
            <%else%>
            <%=jj%>
            <%end if %>
      </span><br>
        <hr color="#FFFFFF"></hr>        </td>
            </tr>
    <% rsnei.movenext
		if rsnei.eof then exit for 
		next %>
    </table>
          <table width="675" height="31" border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td height="31">                 　　　　　　　　　　　　　　　　　<div align="right" class="STYLE6">
                <% if page<>1 then %>
                <a href=<%=path%>?page=1&page1=<%=clng(page1)%>> 第一页</a> <a href=<%=path%>?page=<%=(page-1)%>&page1=<%=clng(page1)%>> 上一页</a>
                <%end if 
		     if page<>rsnei.pagecount then %>
                <a href=<%=path%>?page=<%=(page+1)%>&page1=<%=clng(page1)%>>下一页</a> <a href=<%=path%>?page=<%=rsnei.pagecount%>&page1=<%=clng(page1)%>>最后一页</a>
                <%end if %>
                </div></td>
              </tr>
            </table>
            <table width="675" height="21" border="0">
              <tr>
                <td width="563"><img src="images/FGline.gif"></td>
              </tr>
            </table>          
     <table width="675" height="81" border="0" cellpadding="0" cellspacing="0" style="color:#336699">
       <tr>
         <td height="25"><div align="center" class="style3"><strong>国外IT企业故事</strong></div></td>
              </tr>
       <tr>
         <td height="20"><div align="center"></div></td>
              </tr>
       <% '分页'
	  rswai.pagesize=6
	  page1=clng(request("page1"))
	  if page1<1 then page1=1
	  rswai.absolutepage=page1
	  for i=1 to rswai.pagesize
	  jjj=rswai("introduce")
  %>
       <tr>
         <td height="45" ><span class="STYLE6"><a href="#"onClick="javascript:window.open('introduce7.asp?id=<%=rswai("ID")%>','','width=650,height=550')"><%=rswai("name1")%></a><br>
               <% if len(jjj)>200 then%> 
               <%=left(jjj,200)&"...."%>
               <%else%>
               <%=jjj%>
               <%end if %>
         </span><br>
           <hr color="#FFFFFF"></hr>           </td>
              </tr>
       <% rswai.movenext
		if rswai.eof then exit for 
		next %>
       </table>
            <table width="675" height="18" border="0" cellpadding="0" cellspacing="0"style="color:#336699">
              <tr>
                <td width="622" height="18">
                  <div align="right" class="STYLE6">
                    <% if page1<>1 then %>
                    <a href=<%=path%>?page1=1&page=<%=clng(page)%>>第一页</a> <a href=<%=path%>?page1=<%=(page1-1)%>&page=<%=clng(page)%>>上一页</a>
                    <%end if %>
                    <% if page1<>rswai.pagecount then %>
                    <a href=<%=path%>?page1=<%=(page1+1)%>&page=<%=clng(page)%>>下一页</a> <a href=<%=path%>?page1=<%=rswai.pagecount%>&page=<%=clng(page)%>>最后一页</a>
                    <%end if %> 
                  </div></td>
              </tr>
          </table></tr>
    </table>
   </td> 
   </tr> 
</table> 
</body>
</html>
