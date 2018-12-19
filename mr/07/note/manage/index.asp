<!--#include file="../common/conn.asp"-->
<!--#include file="../Replace.asp"-->
<%
Set rs = Server.CreateObject("ADODB.RecordSet")
//检索数据信息
sql= "select tb_note.*,answ.* from tb_note left join (select noan_note_id from tb_note_answer group by noan_note_id) as answ  on answ.noan_note_id = tb_note.note_id  where note_flag=0 order by note_time desc"
	rs.Open sql,conn,1,3
	rs.PageSize = 10
	'实现分页
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
//反选表单中checkbox
function inverse(form){
	for (var i=0;i<form.elements.length-2;i++){
		var e = form.elements[i];
		e.checked == true ? e.checked = false : e.checked = true;
	}
}
//选择表单中所有check_box
function check_all(form){
	for(var i=0;i<form.elements.length-2;i++){
		var e=form.elements[i];
		e.checked=true;
	}
}
</script>
<html>
<head>
<title>美丽心情留言本后台管理</title>
<meta http-equiv="Content-Type" content="text/html;charset=gb2312">
<link rel="stylesheet" type="text/css" href="../css/css.css">
<style type="text/css">
<!--
.STYLE22 {color: #CC3333}
-->
</style>
</head>
<table width="950" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td colspan="3"><!--#include file="m_header.html"--></td>
  </tr>
  <tr>
    <td width="190" valign="top" background="../images/xline.jpg" bgcolor="#F7F7F7"><!--#include file="m_left.asp"--></td>
    <td width="757" valign="top" bgcolor="#FFFFFF"><table width="757" cellpadding="0" cellspacing="0" border="0" bgcolor="#ffffff">
        <tr align="left">
          <td height="50" align="right" bgcolor="#D9EBB9" class="border"><table width="100%" align="center" border="0" cellspacing="0" cellpadding="0" >
              <tr height="24">
                <td width="765" height="52" background="../images/right.jpg">&nbsp;</td>
              </tr>
            </table>
            <table width="757" border="0" align="center" cellpadding="0" cellspacing="0" class="tb1">
              <form action="./note_delete.asp" method="post" name="note" id="note">
                <tr align="left">
                  <td class="border"></td>
                </tr>
                <tr class='style1'>
                  <td width='58' height="26" align="center" bgcolor="D9EBB9">选择</td>
                  <td width='275' align="center" bgcolor="D9EBB9">标题</td>
                  <td colspan="2" align="center" bgcolor="D9EBB9">作者</td>
                  <td width='140' align="center" bgcolor="D9EBB9">是否回复</td>
                  <td width='135' align="center" bgcolor="D9EBB9">发布时间</td>
                </tr>
				<%
				repeat_rows = 0
				While(repeat_rows < rs.PageSize and not rs.Eof)
				title=rs("note_title")
				author=rs("note_user")
				time1=rs("note_time")
				%>
                <tr valign="middle" height="24">
                  <td align="center" bgcolor="#F8F7EB"><input type="checkbox" name="note_id"  value="<% =rs("note_id") %>" /></td>
                  <td bgcolor="#F8F7EB"><!-- 在标题上添加连接 -->
                    <a href="note_read.asp?note_id=<% =rs("note_id")%>" target="_blank">
                  <% response.Write(title) %></a></td>
                  <!-- 作者 -->
                  <td colspan="2" align="center" bgcolor="#F8F7EB"><% response.Write(author) %></td>
                  <td align="center" bgcolor="#F8F7EB">
				  <%
				  if rs("note_answer")=0 then
				  	 	response.Write("未回复")
				  else
						response.Write("已回复")
				  end if
				  %>					</td>
                  <td align="center" bgcolor="#F8F7EB"><% response.Write(time1)%></td>
                </tr>
				<%
				repeat_rows = repeat_rows + 1
				rs.MoveNext
				Wend
				%>
                <tr align="right" bgcolor="#F0F0F0" valign="middle">
                  <td height="28" colspan="6" align="left" bgcolor="F8F7EB"> &nbsp;&nbsp;
                    <input type="checkbox" name="select_all" onClick="check_all(this.form)" />
                    <span class="style1">全选</span>&nbsp;
                    <input type="checkbox" name="select_reverse" onClick="inverse(this.form)" />
                    <span class="style1">反选</span>
                    <input name="submit" type="submit" value = " 删除 " class="btn1" />
                    &nbsp;&nbsp; 每页『<span class="STYLE22">&nbsp;<%=rs.PageSize%>&nbsp;</span>』条/共『<span class="STYLE22">&nbsp;<%=rs.RecordCount%>&nbsp;</span>』条&nbsp;当前『&nbsp;<span class="STYLE22"><%=pageno%>&nbsp;</span>』页/共『<span class="STYLE22">&nbsp;<%=rs.PageCount%>&nbsp;</span>』页&nbsp;
                    <%
					if(pageno <> 1) then
					%>
					&nbsp;<a href="index.asp?pageno=1">第一页</a>
					<%
					End if
					if(pageno <> 1)then
					%>
					<a href="index.asp?pageno=<%=(pageno-1)%>">上一页</a>
					<%
					end if
					if(instr(pageno,cstr(rs.pagecount)) = 0) then
					%>
					<a href="index.asp?pageno=<%=(pageno+1)%>">下一页</a>
					<%
					end if	
					if(instr(pageno,cstr(rs.pagecount)) = 0) then
					%>
					<a href="index.asp?pageno=<%=rs.pagecount%>">最后一页</a>
					<%
					end if
					rs.close
					Set rs = Nothing
					%></td>
                </tr>
              </form>
            </table>
           </td>
        </tr>
    </table>
    </td>
  </tr>
  <tr>
    <td colspan="3"><!--#include file="footer.html"--></td>
  </tr>
</table>
