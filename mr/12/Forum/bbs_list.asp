<!--#include file="Conn/conn.asp" --><!--�������ݿ������ļ�-->
<%
typeid=request.QueryString("TypeID")
if typeid<>"" then								'����������Ϣ��ID
set rs1=server.CreateObject("ADODB.RecordSet")						'������¼��
sql1="select * from view_smalltype where typeid="&session("id")		'��ѯ����
rs1.open sql1,conn,1,3												'�򿪼�¼��
end if
%>
<script src="JS/fun.js"></script>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
.STYLE2 {font-size: 9pt}
a:link {
	text-decoration: none;
}
a:visited {
	text-decoration: none;
}
a:hover {
	text-decoration: none;
}
a:active {
	text-decoration: none;
}
-->
</style>
<script language="javascript">
function mycheck(){
if(findbbs.keyword.value==""){alert("��������Ҫ��ѯ�Ĺؼ���!");findbbs.keyword.focus();return false}
	findbbs.submit();
}
</script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="css/style.css">
<!--#include file="top.asp"-->
<table width="96%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%" height="23" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
      <tr>
        <td bgcolor="#E9EBEB">
		<% if typeid<>"" then %>
		<table width="100%" height="22" border="0" align="center" cellpadding="0" cellspacing="0">
         <form name="findbbs" method="post" action="bbs_chaxun.asp" onsubmit="return mycheck(this)">
		  <tr>
            <td width="3%"><div align="right"><img src="images/bbs_biao.gif" width="16" height="14"></div></td>
            <td width="61%">&nbsp;<span class="STYLE2">��̴ʵ似����̳&nbsp;��&nbsp;<%=rs1("typename")%></span></td>
            <td width="20%"><input type="text" name="keyword" size="30" class="inputcss1"></td>
            <td width="1%">&nbsp;</td>
            <td width="15%"><input type="image"  src="images/findbbs_button.gif" alt="��������"></td>
          </tr>
		  </form>
        </table>
		<% end if %>
		</td>
      </tr>
    </table></td>
  </tr>
</table>
<table width="96%" height="5" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td bgcolor="#FFFFFF"></td>
  </tr>
</table>
<table width="94%" height="25" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="75%"><table  height="25" border="0" align="left" cellpadding="0" cellspacing="0">
      <tr>
        <td width="65" height="18" background="{$fontallimage}" valign="bottom"style="cursor:hand"></td>
        <td width="3" height="10"></td>
        <td width="65" height="18" background="{$fonthotimage}" valign="bottom" style="cursor:hand"></td>
        
        <td width="3" height="10"></td>
        <td width="65" height="18" background="{$fontundoimage}" valign="bottom" style="cursor:hand"></td>
		
		
		
		<td width="3" height="10"></td>
        <td width="65" height="18" background="{$fontnowdoimage}" valign="bottom" style="cursor:hand"></td>
		<td width="3" height="10"></td>
        <td width="65" height="18" background="{$fontdooverimage}" valign="bottom" style="cursor:hand"></td>
        
        <td width="3" height="10"></td>
      </tr>
    </table></td>
    <td width="7%"><a href="yanzheng.asp?typeid=<%=typeid%>"><img src="images/pubbbs_button.gif" width="60" height="22" border="0" align="absbottom"></a>	</td>
    <td width="1%">&nbsp;</td>
    <td width="17%"><img src="images/refreash_button.gif" width="60" height="22" style="cursor:hand" onClick="location.reload()"></td>
  </tr>
</table>
<table width="96%" height="6" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td background="images/line_1.gif"></td>
  </tr>
</table>
<table width="96%" height="3" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td bgcolor="#FFFFFF"></td>
  </tr>
</table>

<%
If Session("UserName")<>"" Then
	UStatus=""
	Set rs=Server.CreateObject("ADODB.RecordSet")
	SQL="Select * From tb_User Where UserName='"&Session("UserName")&"'"
	rs.Open SQL,Conn,1,3	
	If rs.Eof and rs.Bof Then
		Session("UserName")=""
		UStatus=""
	Else
		If rs("Status")="����" Then 	UStatus="����" End If
	End If
End If
if TypeID<>"" then 
'��ѯ������Ϣ
Set rs=Server.CreateObject("ADODB.RecordSet")
SQL="Select * from view_smalltype where ID="&TypeID
rs.Open SQL,Conn,1,3
smalltypename=rs("smalltypename")
'session("id")=rs("id")
Set rs=Server.CreateObject("ADODB.RecordSet")
SQL="Select * From view_board where TypeID="& TypeID &" and state='��ͨ����' Order By CreateTime DESC"

rs.Open SQL,Conn,1,3
Set rs_B=Server.CreateObject("ADODB.RecordSet")
SQL="Select * From view_board where TypeID="& TypeID &" and state='��������' Order By CreateTime DESC"
rs_B.Open SQL,Conn,1,3
else
	response.Redirect("index.asp")
End If
%>
<%
'ת��UBB����
Function UBBCodeDeal(InString)
	InString1=Server.HTMLEncode(InString)
	InString1=Replace(InString1,"[","<")
	InString1=Replace(InString1,chr(13),"<BR>")
	UBBCodeDeal=Replace(InString1,"]",">")
End Function
%> 
<table width="777" height="341"  border="0" align="center" cellpadding="-2" cellspacing="-2" bgcolor="#FFFFFF">
  <tr>
    <td valign="top"><table width="96%" border="0" align="center" cellpadding="-2" cellspacing="-2" class="tableBorder">
      <tr>
        <td></td>
      </tr>
      <tr>
        <td></td>
      </tr>
      <tr>
        <td height="2"></td>
      </tr>
      <tr>
        <td valign="top">
      <table width="100%" height="205"  border="0" cellpadding="-2" cellspacing="-2" class="tableBorder">
        <tr>
          <td width="13" height="32" background="Images/bg.gif">&nbsp;</td>
          <td width="689" background="Images/bg.gif" style="color:#FFFFFF"> �� <%=smalltypename%> ר��


 �� </td>
          <td width="73" background="Images/bg.gif">&nbsp;</td>
        </tr>
        <tr valign="top">
          <td height="171" colspan="3"><table width="100%" height="129"  border="0" cellpadding="-2" cellspacing="-2">
            <tr>
              <td align="center" valign="top" bgcolor="FBFBFB">
			  <%
			  If rs.Eof and rs.Bof and rs_B.eof and rs_B.bof Then%>
					<table width="100%" height="152"  border="0" cellpadding="-2" cellspacing="-2">
                      <tr>
                        <td width="15%" align="center" bgcolor="FBFBFB"><img src="Images/clue.gif" width="50" height="57"></td>
                        <td width="82%" bgcolor="FBFBFB">�� <%=smalltypename%> ��ר������������Ϣ���������������⡱���������ӣ�</td>
                        <td width="3%" rowspan="2" bgcolor="FBFBFB">&nbsp;</td>
                      </tr>
                      <tr>
                        <td height="34" colspan="2" align="center" bgcolor="FBFBFB"><div align="center">[ <a href="#" onClick="window.location.href='main.asp';">�鿴�������</a> ]</div></td>
                      </tr>
                    </table>
	 <%Else%>	
			  <table width="100%" height="46"  border="1" cellpadding="-2" cellspacing="-2" class="tableBorder_LR" bordercolor="#FFFFFF" bordercolordark="#9999CC" bordercolorlight="#FFFFFF">
			  <tr align="center" height="25">
			  <td width="38" height="5" bgcolor="88BAF9">״̬</td>
			  <td width="34" bgcolor="88BAF9">����</td>
			  <td width="416" bgcolor="88BAF9">���� [���+����չ�������б�]</td>
			  <td width="88" bgcolor="88BAF9">����</td>
			  <td width="65" bgcolor="88BAF9">�ظ�/����</td>
			  <td width="134" bgcolor="88BAF9">����ʱ��</td>
			  </tr>
<%' *****��ʾ����������Ϣ*****
			  m=0
			if not(rs_B.eof and rs_B.bof) then
			while not (rs_B.eof)
			  %>
			  <tr height="27">
			    <td align="center"><img src="Images/bbs.gif" alt="��������" width="20" height="16"></td>
                <td align="center"><img src="Images/Face/<%=rs_B("face")%>" width="20" height="20"></td>
			    <td>
					<%
					If(rs_B("reply")=0)Then
					%>
					&nbsp;
					<%
					Else
					%>
					&nbsp;<a href="Javascript:ShowTR(img<%=m%>,OpenRep<%=m%>)"></a>
					<%
					End If
					%>					<a href="show.asp?TopicID=<%=rs_B("ID")%>"><%=server.HTMLEncode(rs_B("subject"))%></a>                     </td>
			    <td align="center">&nbsp;<%=rs_B("username")%></td>
			    <td align="center">&nbsp;<%=rs_B("reply")%>/<%=rs_B("hit")%></td>
			    <td align="center">&nbsp;<%=rs_B("createTime")%></td>
			  </tr>
			<%
			If(rs_B("reply")>0)Then
				set rs_reply=Server.CreateObject("ADODB.RecordSet")
				sql="select * from view_reply where TopicID="&rs_B("ID")&" order by CreateTime desc"
				rs_reply.open sql,conn,1,3
				%>
			  <tr id="OpenRep<%=m%>" style="display:none;">
			   <td colspan="6">				
				<% for j=1 to rs_reply.recordcount%>
			     <table width="100%"  border="0" cellspacing="-2" cellpadding="-2" bgcolor="#F9F9FF">
                   <tr onMouseOver="this.style.background='#EFEFFF'" onMouseOut="this.style.background=''">
                      <td width="46" height="18">&nbsp;</td>				   
                      <td width="92%">&nbsp;<a href="show.asp?TopicID=<%=rs_B("ID")%>"><%=server.HTMLEncode(rs_reply("subject"))%> &nbsp;&nbsp;
					</a></td>
                   </tr>
                 </table>
				 <%	m=m+1  'ע��:�������һ��������
			  		rs_reply.movenext
				next
				%></td>
			  </tr>
			  <%
			  End if
			  rs_B.MoveNext
			  wend
			  rs_Bnum=rs_B.recordcount
		      end if
%>

 
<%' ***��ʾ��ͨ������Ϣ***
	if not(rs.eof and rs.bof) then
			'��ҳ��ʾ����б�
			rs.pagesize=10
			page=CLng(Request("page"))
			if page<1 then page=1
			rs.absolutepage=page
			for i=1 to rs.pagesize
			  %>
			  <tr height="27">
			    <td align="center"><img src="Images/reader.gif" alt="��ͨ����" width="21" height="15"></td>
                <td align="center"><img src="Images/Face/<%=rs("face")%>" width="20" height="20"></td>
<script language="javascript">
function ShowTR(objImg,objTr){
	if(objTr.style.display == "block"){
		objTr.style.display = "none";
		objImg.src = "Images/jia.gif";
		objImg.alt = "չ��";		
	}else{
		objTr.style.display = "block";
		objImg.src = "Images/jian.gif";
		objImg.alt = "�۵�";		
	}
}
</script>



			    <td id="Nq<%=m%>">
					<%
					If(rs("reply")=0)Then
					%>
					&nbsp;<img src="Images/jian.gif" border="0">
					<%
					Else
					%>
					&nbsp;<a href="Javascript:ShowTR(img<%=m%>,OpenRep<%=m%>)"><img src="Images/jia.gif" border="0" alt="չ��" id="img<%=m%>"></a>
					<%
					End If
					%>					<a href="show.asp?TopicID=<%=rs("ID")%>"><%=server.HTMLEncode(rs("subject"))%></a>                     </td>
			    <td align="center">&nbsp;<%=rs("username")%></td>
			    <td align="center">&nbsp;<%=rs("reply")%>/<%=rs("hit")%></td>
			    <td align="center">&nbsp;<%=rs("createTime")%></td>
			  </tr>
			<%
			If(rs("reply")>0)Then
				set rs_reply=Server.CreateObject("ADODB.RecordSet")
				sql="select * from view_reply where TopicID="&rs("ID")&" order by CreateTime asc"
				rs_reply.open sql,conn,1,3
				%>
			  <tr id="OpenRep<%=m%>" style="display:none;">
			   <td colspan="6">				
				<% for j=1 to rs_reply.recordcount%>
			     <table width="100%"  border="0" cellspacing="-2" cellpadding="-2" bgcolor="#F9F9FF">
                   <tr onMouseOver="this.style.background='#EFEFFF'" onMouseOut="this.style.background=''">
                      <td width="46">&nbsp;</td>				   
                      <td width="92%">&nbsp;<a href="show.asp?TopicID=<%=rs("ID")%>#<%response.Write(rs_reply("ID")) '��ת��show.aspҳ��ê�㴦%>"><%=server.HTMLEncode(rs_reply("subject"))%> &nbsp;&nbsp;
					  <font class="word_grey">[�ظ�ʱ�䣺<%=rs_reply("createTime")%>]</font></a></td>
                   </tr>
                 </table>
				 <%	m=m+1  'ע�⣬�������һ��������
			  		rs_reply.movenext
				next
				%>			   </td>
			  </tr>
			  <%
			  End if
			  rs.MoveNext
			  If rs.Eof Then exit For 
			  Next
			  end if
%>
              </table>			</td>
            </tr>
          </table>		  </td>
        </tr>
      </table>
	  
<table width="100%" height="25"  border="0" cellpadding="-2" cellspacing="-2">
  <tr>
    <td valign="bottom"><% if page<>1 then %><a href=?page=1&TypeID=<%=TypeID%>>��һҳ</a>
				<a href=?page=<%=(page-1)%>&TypeID=<%=TypeID%>>��һҳ</a> 
		    <%end if 
			if page<>rs.pagecount then %>
   				<a href=?page=<%=(page+1)%>&TypeID=<%=TypeID%> >��һҳ</a> 
				<a href=?page=<%=rs.pagecount%>&TypeID=<%=TypeID%>>���һҳ</a> 
          <%end if%>	</td>
	<%If not(rs.Eof and rs.Bof) Then%>	
    <td width="29%" align="right" class="word_grey">[<%=page%>/<%=rs.PageCount%>]&nbsp;&nbsp;ÿҳ<%=rs.PageSize%>��&nbsp;&nbsp;��<%=rs.RecordCount%>��������Ϣ<%if rs_Bnum>0 then%>,<%=rs_Bnum%>����������<%end if%></td>
		<%End If 
End If
	%>	
  </tr>
</table>    </td>
  </tr>
  <tr>
    <td valign="top" background="Images/hr.gif"></td>
  </tr>
</table>
</body>
</html>
<%
rs.close()
Set rs=Nothing
%>


