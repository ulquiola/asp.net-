<!--#include file="../common/conn.asp"-->
<!--#include file="../Replace.asp"-->
<%
Set rs = Server.CreateObject("ADODB.RecordSet")
//����������Ϣ
sql = "select * from tb_note where note_flag=1 order by note_time desc"
rs.Open sql,conn,1,3
rs.PageSize = 10
'ʵ�ַ�ҳ
	if rs.eof then
		rs_total = 0
	else
		rs_total = rs.RecordCount
	end if
	dim pageno
	getpageno = Switch(Request("pageno"))
	if(getpageno = "") then
		pageno = 1
	else
		pageno = getpageno
	end if
	if(not rs.Eof) then
		rs.AbsolutePage = pageno
	end if
%>
<script language="JavaScript">
<!--
//��ѡ����checkbox
function inverse(form){
	for (var i=0;i<form.elements.length-2;i++){
		var e = form.elements[i];
		e.checked == true ? e.checked = false : e.checked = true;
	}
}
//ѡ���������check_box
function check_all(form){
	for(var i=0;i<form.elements.length-2;i++){
		var e=form.elements[i];
		e.checked=true;
	}
}
//-->
</script>
<html>
<head>
<title>�����������Ա�</title>
<meta http-equiv="Content-Type" content="text/html;charset=gb2312">
<link rel="stylesheet" type="text/css" href="../css/css.css">
</head>
<body>
<table width="950" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td colspan="3"><!--#include file="m_header.html"--></td>
  </tr>
  <tr>
    <td width="190" valign="top" bgcolor="#EEEEEE"><!--#include file="m_left.asp"--></td>
    <td width="757" valign="top" bgcolor="#FFFFFF"><table width="757" cellpadding="0" cellspacing="0" border="0" bgcolor="#ffffff">
        <tr align="left">
          <td class="border"><table width="100%" align="center" border="0" cellspacing="0" cellpadding="0" >
              <tr height="24">
                <td height="52" background="../images/right.jpg">&nbsp;</td>
              </tr>
            </table>
            <table width="757" border="0" align="center" cellpadding="0" cellspacing="0" class="tb1">
              <form action="./m_note_del.asp" method="post" name="note" id="note">
                <tr align="left">
                  <td class="border"></td>
                </tr>
                <tr class='style1'>
                  <td width='45' height="26" align="center" bgcolor="D9EBB9">ѡ��</td>
                  <td width='338' align="center" bgcolor="D9EBB9">����</td>
                  <td width='195' align="center" bgcolor="D9EBB9">����</td>
                  <td width='179' align="center" bgcolor="D9EBB9">����ʱ��</td>
                </tr>
				<%
				repeat_rows = 0
				While(repeat_rows < rs.PageSize and not rs.Eof)
				id=rs("note_id")
				title=rs("note_title")
				author=rs("note_user")
				time1=rs("note_time")
				note_flag=rs("note_flag")
				%>
                <tr valign="middle" height="24">
                  <td align="center"><input type="checkbox" name="note_id"	value="<% response.Write(id) %>" /></td>
                  <td><!-- �ڱ������������ -->
                    <a href="m_note_read.asp?note_id=<% response.Write(id)%>" target="_blank"> <% response.Write(title) %>
                    <!-- ���� -->
                    </a></td>
                  <!-- ���� -->
                  <td><% response.Write(author) %></td>
                  <td><% response.Write(time1) %></td>
                </tr>
                <%
				repeat_rows = repeat_rows + 1
				rs.MoveNext
				Wend
				if (rs.bof and rs.eof)then
				%>
				 <tr valign="middle"  height="30">
                   <td height="25" colspan="5" align="center" bgcolor="#F0F0F0">  û���˶���˵���Ļ���</td>
                </tr>
				<%
				end if
				%>
                <tr align="right" bgcolor="#F0F0F0" valign="middle">
                  <td height="28" colspan="6" align="left" bgcolor="F8F7EB"> &nbsp;
                    <input type="checkbox" name="select_all" onClick="check_all(this.form)" />
                    <span class="style1">ȫѡ</span>&nbsp;
                    <input type="checkbox" name="select_reverse" onClick="inverse(this.form)" />
                    <span class="style1">��ѡ</span>
                    <input name="submit" type="submit" value = " ɾ�� " class="btn1" />
                    &nbsp;&nbsp; ÿҳ��<span class="STYLE22">&nbsp;<%=rs.PageSize%>&nbsp;</span>����/����<span class="STYLE22">&nbsp;<%=rs.RecordCount%>&nbsp;</span>����&nbsp;��ǰ��&nbsp;<span class="STYLE22"><%=pageno%>&nbsp;</span>��ҳ/����<span class="STYLE22">&nbsp;<%=rs.PageCount%>&nbsp;</span>��ҳ&nbsp;
                    <%
					if(pageno <> 1) then
					%>
					&nbsp;<a href="manage_note.asp?pageno=1">��һҳ</a>
					<%
					End if
					if(pageno <> 1)then
					%>
					<a href="manage_note.asp?pageno=<%=(pageno-1)%>">��һҳ</a>
					<%
					end if
					if(instr(pageno,cstr(rs.pagecount)) = 0) then
					%>
					<a href="manage_note.asp?pageno=<%=(pageno+1)%>">��һҳ</a>
					<%
					end if	
					if(instr(pageno,cstr(rs.pagecount)) = 0) then
					%>
					<a href="manage_note.asp?pageno=<%=rs.pagecount%>">���һҳ</a>
				  <%
					end if
					rs.close
					Set rs = Nothing
					%></td>
                </tr>
                <tr>
                  <td ></td>
                  <td colspan="4" height="5"></td>
                </tr>
              </form>
            </table></td>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td colspan="3"><!--#include file="footer.html"--></td>
  </tr>
</table>
