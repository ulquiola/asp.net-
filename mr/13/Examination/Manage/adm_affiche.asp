<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Conn/conn.asp"-->
<!--#include file="../inc/customFunc.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>������Ϣ����</title>
<link href="../Css/style.css" rel="stylesheet">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
}
-->
</style>
</head>
<body> 
<table width="95%"  border="0" cellspacing="0" cellpadding="0"> 
  <tr> 
    <td>&nbsp;</td> 
  </tr> 
</table> 
<div align="right">[<a href="adm_affiche.asp?action=add"><font color="#FF0000">��ӹ�����Ϣ</font></a>]</div>
<%

	  Dim action
	  action=Request.QueryString("action")
	  Select case action
	  case "add"
	    call add()	
	  case "savenew"  
	    call savenew()
	  case "edit"
	    call edit()
	  case "saveedit"
	    call saveedit()
	  case "del"
	    call del()
	  End Select
	%> 
<%If action="" Then%> 
<table width="95%"  border="0" cellspacing="1" cellpadding="1"> 
  <form name="form12" method="post" action=""> 
    <tr> 
      <td width="150" height="22" align="right">��ѯ�ؼ���:</td> 
      <td width="200" height="22"><input name="keyword" type="text" id="keyword" size="30"></td> 
      <td width="100" height="22"><select name="sel_item" id="sel_item"> 
          <option value="1">�������</option> 
        </select></td> 
      <td width="120" height="22"><input type="submit" name="Submit" value="�� ѯ"></td> 
    </tr> 
  </form> 
</table> 
<table width="98%"  border="1" cellspacing="0" cellpadding="0" align="center"
 bordercolor="#FFFFFF" bordercolordark="#1985DB" bordercolorlight="#FFFFFF"> 
  <tr> 
    <td width="254" height="20" align="center" valign="middle">�������</td> 
    <td width="311" height="20" align="center" valign="middle">����ʱ��</td> 
    <td width="159" height="20" align="center" valign="middle">����</td> 
  </tr> 
  <script language="javascript">
		function ShowTR(objImg,objTr){
		  if(objTr.style.display == "block"){
			objTr.style.display = "none";
			objImg.src = "images/jia.gif";
			objImg.alt = "չ��";		
		  }else{
			objTr.style.display = "block";
			objImg.src = "images/jian.gif";
			objImg.alt = "�۵�";		
		  }
		}
	 </script> 
  <%	  
	  keyword=Request("keyword")
	  sel_item=Request("sel_item")
	  Set rs=Server.CreateObject("ADODB.Recordset")
	  sqlstr="select * from tb_affiche where 1=1"
	  If sel_item<>"" and keyword<>"" Then
		If sel_item=1 Then sqlstr=sqlstr&" and Aftitle like '%"&keyword&"%'"
	  End If
	  sqlstr=sqlstr&" order by id desc"		  
	  rs.open sqlstr,conn,1,1
	  If rs.eof or rs.bof Then		
		Response.Write("<tr align='center'><td colspan='3'>��ʱ�޹�����Ϣ!<td><tr>")    
	  Else
	  rs.pagesize=6  '����ÿҳ��ʾ�ļ�¼��
	  pages=clng(Request("pages"))  '��õ�ǰҳ��
	  If pages<1 Then pages=1
	  If pages>rs.recordcount Then pages=rs.recordcount
	  showpage rs,pages  'ִ�з�ҳ�ӳ���showpage		
	  Sub showpage(rs,pages)  '��ҳ�ӳ���showpage(rs,pages)
	  Set rs_account=Server.CreateObject("ADODB.Recordset")
	  rs.absolutepage=pages   'ָ��ָ�����ڵĵ�ǰλ��
	  For i=1 to rs.pagesize  'ѭ����ʾ��¼���еļ�¼
	  %> 
  <tr> 
    <td width="254" height="20" style="text-indent:10px"><a href="Javascript:ShowTR(img<%=rs("id")%>,OpenRep<%=rs("id")%>)"><img src="images/jia.gif" border="0" alt="չ��" id="img<%=rs("id")%>"> <%=rs("Aftitle")%></a></td> 
    <td width="311" height="20" align="center"><%=rs("Afdate")%></td> 
    <td width="159" height="20" align="center">[<a href="?action=edit&editid=<%=rs("id")%>&edittitle=<%=rs("Aftitle")%>&editcontent=<%=rs("Afcontent")%>">�޸�</a>]&nbsp;[<a href="?action=del&delid=<%=rs("id")%>">ɾ��</a>]</td> 
  </tr> 
  <tr align="center" bgcolor="#FFFFFF" id="OpenRep<%=rs("id")%>" style="display:none;"> 
    <td height="20" colspan="3"><table width="95%"  border="0" cellspacing="1" cellpadding="1"> 
        <tr> 
          <td align="left" style="line-height:18px"><%=rs("Afcontent")%></td> 
        </tr> 
      </table></td> 
  </tr> 
  <% 
		rs.movenext  'ָ�������ƶ�
		If rs.eof Then exit for
		Next
		End Sub
		End If	
	  %> 
  <tr> 
    <td height="20" colspan="4"><table width="100%" height="25"  border="0" cellpadding="-2" cellspacing="-2"> 
        <tr> 
          <%If Not (rs.eof and rs.bof) Then%> 
          <td width="62%" valign="bottom"><% if pages<>1 then %> 
            <a href=?pages=1&keyword=<%=keyword%>&sel_item=<%=sel_item%>>��һҳ</a> <a href=?pages=<%=(pages-1)%>&keyword=<%=keyword%>&sel_item=<%=sel_item%>>��һҳ</a> 
            <%end if 
			if pages<>rs.pagecount then %> 
            <a href=?pages=<%=(pages+1)%>&keyword=<%=keyword%>&sel_item=<%=sel_item%>>��һҳ</a> <a href=?pages=<%=rs.pagecount%>&keyword=<%=keyword%>&sel_item=<%=sel_item%>>���һҳ</a> 
            <%end if%> </td> 
          <%If not(rs.Eof and rs.Bof) Then%> 
          <td width="38%" align="right" class="word_grey">[<%=pages%>/<%=rs.PageCount%>]&nbsp;&nbsp;[��<font color="red"><%=rs.RecordCount%></font>����Ϣ]</td> 
          <%End If 
	  Set rs_account=Nothing
	  rs.close
	  Set rs=Nothing
	 End If
	%> 
        </tr> 
      </table></td> 
  </tr> 
</table> 
<%End If%> 
<%Sub add()%> 
<table width="95%"  border="0" cellspacing="1" cellpadding="1"> 
  <form name="form1" method="post" action="?action=savenew"> 
    <tr> 
      <td height="22" colspan="2" align="center">��ӹ�����Ϣ</td> 
    </tr> 
    <tr> 
      <td width="150" height="22" align="right">�������:</td> 
      <td height="22" align="left"><input name="txt_name" type="text" id="txt_name" size="40"></td> 
    </tr> 
    <tr> 
      <td width="150" height="22" align="right">��������:</td> 
      <td height="22" align="left"><textarea name="txt_content" cols="50" rows="5" id="txt_content"></textarea></td> 
    </tr> 
    <tr> 
      <td width="150" height="22" align="right">&nbsp;</td> 
      <td height="32" align="left"><input type="submit" name="Submit" value="  �ύ  "> 
&nbsp; 
        <input type="reset" name="Submit2" value="  ����  ">
       &nbsp; <input type="button" name="Submit3" value="  ����  " onClick="javascript:history.back();"></td> 
    </tr> 
  </form> 
</table> 
<% End Sub %>
<%	Sub savenew()
	  txt_name=filter_Str(Request.Form("txt_name"))
	  txt_content=filter_Str(Request.Form("txt_content"))
	  If txt_name="" or txt_content="" Then
	    Response.Write("<script language='javascript'>alert('�������������ݲ���Ϊ��,����������!');history.back();</script>")
	  Else
	    Set rs=Server.CreateObject("ADODB.Recordset")
		sqlstr="select * from tb_affiche where id is null"
		rs.open sqlstr,conn,1,3
		rs.addnew
		rs("Aftitle")=txt_name
		rs("Afcontent")=txt_content
		rs("Afdate")=now()
		rs.update
		rs.close
		Set rs=Nothing
		Response.Redirect("adm_affiche.asp")
	  End If
	  End Sub
	  Sub del()
	  delid=Request.QueryString("delid")
	  conn.Execute("delete from tb_affiche where id="&delid&"")
	  Response.Redirect("adm_affiche.asp")
	  End Sub
	 %>
<% Sub edit() %>
<table width="95%"  border="0" cellspacing="1" cellpadding="1"> 
  <form name="form1" method="post" action="?action=saveedit"> 
    <tr> 
      <td height="22" colspan="2" align="center">�޸Ĺ�����Ϣ</td> 
    </tr> 
    <tr> 
      <td width="150" height="22" align="right">�������:</td> 
      <td height="22" align="left"><input name="txt_name" type="text" id="txt_name" size="40" value="<%=Trim(Request.QueryString("edittitle"))%>"></td> 
    </tr> 
    <tr> 
      <td width="150" height="22" align="right">��������:</td> 
      <td height="22" align="left"><textarea name="txt_content" cols="50" rows="5" id="txt_content"><%=Trim(Request.QueryString("editcontent"))%></textarea></td> 
    </tr> 
    <tr> 
      <td width="150" height="22" align="right">&nbsp;</td> 
      <td height="32" align="left"><input type="submit" name="Submit" value="  �ύ  "> 
&nbsp; 
        <input type="reset" name="Submit2" value="  ����  ">
       &nbsp; <input type="button" name="Submit3" value="  ����  " onClick="javascript:history.back();">
       <input name="editid" type="hidden" id="editid" value="<%=Trim(Request.QueryString("editid"))%>"></td> 
    </tr> 
  </form> 
</table>
<% End Sub
Sub saveedit()
  editid=Trim(Request.Form("editid"))
  edittitle=Trim(Request.Form("txt_name"))
  editcontent=Trim(Request.Form("txt_content"))
  If editid="" or edittitle="" or editcontent="" Then
    Response.Write("<script language='JavaScript'>alert('���������Ĺ�����Ϣ��');history.back();</script>")
	Response.End()
  End If 
  Set rs=Server.CreateObject("ADODB.Recordset")
  sqlstr="select * from tb_affiche where Aftitle='"&edittitle&"' and id<>"&editid&""
  rs.open sqlstr,conn,1,1
  If Not (rs.eof and rs.bof) Then
    Response.Write("<script language='JavaScript'>alert('����Ĺ�������ظ������������룡');history.back();</script>")
	Response.End()
  Else
    sqlstr="update tb_affiche set Aftitle='"&edittitle&"',Afcontent='"&editcontent&"' where id="&editid&""
	conn.Execute(sqlstr)
	Response.Write("<script language='JavaScript'>alert('������Ϣ�޸ĳɹ���');location.href='adm_affiche.asp';</script>")
  End IF
End Sub
%> 
</body>
</html>
