<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/conn.asp"--><!--���ݿ������ļ�-->
<!--#include file="Replace.asp"--><!--���ݴ����ļ�-->
<%
Sql = "Select * from tb_news order by time1 Desc"'��ѯ����
Set rs = Server.CreateObject("ADODB.RecordSet")'������¼��
rs.Open Sql,conn,1,3'�򿪼�¼��
rs.PageSize = 9		'����ÿҳ��ʾ��¼����Ŀ
'ʵ�ַ�ҳ
if rs.Eof then		
	rs_total = 0	'����Ĭ��ֵ
else
	rs_total = rs.RecordCount	'ͳ�Ƽ�¼����
end if
dim pageno						'�������
getpageno = Switch(Request("pageno"))	'Ϊgetpageno������ֵ
if(getpageno = "")then					'�ж�getpageno�����Ƿ�Ϊ��
	pageno = 1							'����Ĭ��ֵ
else
	pageno = getpageno					'Ϊpageno������ֵ
End if
if(not rs.Eof)then
	rs.AbsolutePage = pageno
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
<style type="text/css">
<!--
.STYLE3 {font-size: 9pt; color: #000000; }
-->
</style>
</head>
<style type="text/css">
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-bottom: 0px;
}
</style>
<script language="javascript">
function goto(id)
{
	location.href = "Detail1.asp?id=" + id;
}
</script>
<!--#include file="top.asp"-->
<body>
<table width="800" height="529" border="0" align="center" cellpadding="0" cellspacing="0" background="images/new14.gif">
  <tr>
    <td height="86" colspan="3">&nbsp;</td>
  </tr>
  <tr>
    <td width="46" height="289">&nbsp;</td>
    <td width="713" valign="top"><table width="693" height="70" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td width="693" height="70"><table width="662" border="1" align="center"  cellpadding="0" cellspacing="0" bordercolorlight="#F7AB58" bordercolordark="#ffffff">
          <tr>
            <td height="24" align="center" nowrap="nowrap"> ��������</td>
            <td height="25" align="center" nowrap="nowrap"> ��������</td>
            <td align="center" nowrap="nowrap"> ����ʱ��</td>
          </tr>
          <%
	repeat_rows = 0										'Ϊrepeat_rows������ֵ
	While(repeat_rows < rs.PageSize and Not rs.Eof)		
	%>
          <tr style="cursor:hand" onmouseover="this.style.background='#dddddd'" onmouseout="this.style.background='#ffffff'" onclick="goto(<%=rs("Id")%>)">
            <td width="230" height="25" align="center" nowrap="nowrap"><%=Server.HtmlenCode(rs("title"))%> </td>
            <td width="288" align="center" nowrap="nowrap"><%
		If(Len(rs("content")) > 50)Then											'�ж���Ϣ���ݵĳ����Ƿ����50
		Response.Write(Server.HTMLEncode(Mid(rs("content"),1,50))&"....")		'�����Ϣ���ݲ���������
		Else
		Response.Write(Server.HTMLEncode(rs("content")))						'�����Ϣ���ݲ���������
		End If
		%>            </td>
            <td width="167" align="center" nowrap="nowrap"><%=rs("time1")%></td>
          </tr>
          <%
	repeat_rows = repeat_rows + 1												'��repeat_rows������1
	rs.MoveNext																	'�����ƶ���¼ָ��
	Wend
	%>
        </table>
          <table cellspacing="0" cellpadding="0" border="0" height="18" width="662" align="center">
            <tr>
              <td align="left" nowrap="nowrap"> [<%=pageno%>/<%=rs.PageCount%>]&nbsp;ÿҳ<%=rs.PageSize%>��&nbsp;��<%=rs_total%>����¼ </td>
              <td align="right" nowrap="nowrap"><%
					if(pageno <> 1)then
					%>
                  <a href="?">��һҳ</a>
                  <%
					End if
					if(pageno <> 1)then
					%>
                  <a href="?pageno=<%=(pageno-1)%>">��һҳ</a>
                  <%
					end if
					if(instr(pageno,cstr(rs.pagecount)) = 0)then
					%>
                  <a href="?pageno=<%=(pageno+1)%>">��һҳ</a>
                  <%
					end if
					if(instr(pageno,cstr(rs.pagecount)) = 0)then
					%>
                  <a href="?pageno=<%=rs.pagecount%>">���һҳ</a>
                  <%
					end if
					rs.close
					Set rs = Nothing
					%>              </td>
            </tr>
          </table></td>
      </tr>
    </table></td>
    <td width="41">&nbsp;</td>
  </tr>
  <tr>
    <td height="13" colspan="3"><table width="100%" height="21" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td width="40" height="15">&nbsp;</td>
        <td width="19">&nbsp;</td>
        <td width="665">&nbsp;</td>
        <td width="40">&nbsp;</td>
        <td width="34">&nbsp;</td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="133" colspan="3"><table width="800" height="133" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td height="133">&nbsp;</td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
</html>
