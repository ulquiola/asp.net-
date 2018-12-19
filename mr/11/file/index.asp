<!--#include file="Conn.asp"-->
<!--#include file="Function.asp"-->
<!--#include file="inc.asp"-->
<%Call Top("网站公告","index.asp")%>
<table width=785 align="center">
  <tbody>
  <tr>
    <td style="font-size: 12px; width: 200px; height: 21px; text-align: center" 
    align=left rowspan=2 valign="top">
      <table class="table">
        <tbody>
        <tr>
          <td width="200" 
          colspan=3 bgcolor="#3399ff" style="font-weight: bold; width: 200px; color: white; height: 25px; text-align: center; text-decoration: none">用户登录</td>
        </tr>
        <tr>
          <td style="width: 200px; height: 145px; text-align: center" 
          colspan=3 rowspan=2>
            <table style="width: 200px; height: 126px; text-align: left">
              <tbody>
              <tr>
                <td style="width: 188px" colspan=3 rowspan=3><%call login()%></td></tr>
              <tr></tr>
              <tr></tr></tbody></table></td></tr>
        <tr></tr></tbody></table></td>
    <td style="width:100%; height: 21px" align=center valign="top" colspan=2 rowspan=2>
      <table style="width: 100%; height: 182px">
        <tbody>
        <tr>
          <td colspan=3 bgcolor="#3399ff" style="font-weight: bold; color: white; height: 25px; text-align: center; text-decoration: none">网站公告</td>
        </tr>
        <tr>
          <td class="table" colspan=3 rowspan=3 valign="top"><%
		  If Not IsObject(Conn) Then ConnectionDatabase()
		  dim rs
		  set rs=Conn.execute("Select info from tb_Setup")
		  if not rs.eof then
		  response.write htmlencode(rs(0))
		  end if
		  set rs=nothing
		  %></td>
        </tr>
        <tr></tr>
        <tr></tr>
        <tr>
          <td class="titletd" colspan=3>&nbsp;</td>
        </tr></tbody></table></td></tr>
  <tr></tr></tbody></table>
	<%Call Bottom()%>