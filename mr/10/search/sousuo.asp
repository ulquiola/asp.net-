<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn/conn.asp"-->
<!--#include file="conn/function.asp"-->
<%
sousuo=request("sousuo")
prev_sousuo=request("prev_sousuo")
sousuo_within=request("sousuo_within")

if sousuo <> "" then
    if prev_sousuo = "" then 
        prev_sousuo=sousuo
	else
    prev_sousuo=prev_sousuo &","& sousuo
    p_sousuo=split(prev_sousuo,",")
    num_inputted=ubound(p_sousuo)
	end if
sql= "select * from tb_sousuo where content like '%%"& sousuo & "%%'"
	if sousuo_within = "Yes" then
    	for counter =0 to num_inputted-1
        	sql=sql& "and content like '%%"& p_sousuo(counter) & "%%'"
    	next
	end if
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sql,conn,3,3
%>
<script language="javascript">
function Mycheck(){
if (form1.sousuo.value=="")
{ alert("请输入要搜索的全部名称！");form1.sousuo.focus();return ;}
form1.submit();
}
</script>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>搜索引擎</title>
<style type="text/css">
<!--
.STYLE1 {font-size: 9px}
.STYLE2 {font-size: 9pt; }
-->
</style>
</head>
<body onLoad="clockon()" background="images/bg.gif">
<table width="793" border="0" align="center" cellpadding="0" cellspacing="0">
  <!--#include file="shangbiao.asp"-->
  <tr>
    <td width="793" height="588" valign="top"><table width="779" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td height="589" valign="top"  bgcolor="#FFFFFF"><form name="form1" method="post" action="<%= request.servervariables("script_name") %>">
              <table width="780"  border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="63" height="81" align="center" background="images/sousou1.gif"> 　</td>
                  <td width="114" height="81" align="center" background="images/sousou2.gif" class="STYLE2">请输入搜索条件：</td>
                  <td width="214" align="center" background="images/sousou2.gif" class="STYLE2"><input name="sousuo" type="text" id="sousuo" style="height:25px"  value="<%= sousuo %>" size="30"></td>
                  <td width="209" align="left" background="images/sousou2.gif" class="STYLE2"><%
					if sousuo <> "" then %>
                    <input type = "radio" name="sousuo_within" checked value="No">
重新搜索 &nbsp; <input type = "radio" name="sousuo_within" value="Yes">
在结果中搜索
<%
						if sousuo_within = "Yes" then %>
  <input type = "hidden" name="prev_sousuo" value="<%= prev_sousuo %>">
  <%
						else %>
  <input type = "hidden" name="prev_sousuo" value="<%= sousuo %>">
  <% 
						end if%>
  <% end if%></td>
                  <td width="42" height="81" align="center" background="images/sousou2.gif"><input name="imageField" type="image" src="images/souan1.gif" width="42" height="20" border="0" onClick="Mycheck()"></td>
                  <td width="85" height="81" align="center" background="images/sousouhou.gif"><a href="gaojisousuo.asp"><img src="images/souan2.gif" width="77" height="20" border="0"></a></td>
                  <td width="53" align="center" background="images/sousouhou.gif"><a href="index.asp" class="STYLE2"style="color:#3661a6">返回首页</a></td>
                </tr>
                
                <tr>
                  <td height="2" colspan="7"></td>
                </tr>
              </table>
            </form>
            <table width="586" height="66" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
			  <td>&nbsp;</td>
              </tr>
			  <tr>
                <%if rs.eof and rs.bof then%>
                <img src="images/turnbook.gif" width="37" height="30">
                <% response.write "对不起，没有找到您要搜索的信息！"
				   response.End()
				else
				  if sousuo=""then%>
                <img src="images/turnbook.gif" width="37" height="30">
                <% response.write "对不起，没有找到您要搜索的信息！"
				    response.end
				  else
				%>
                <%'分页
				if rs.eof then
					rs_total = 0
				else
					rs_total = rs.recordcount
				end if
				rs.pagesize=5
				page=clng(request("page"))
				if page<1 then page=1
				rs.absolutepage=page
				for i=1 to rs.pagesize
				zzs=rs("content")
				if num_inputted = "Yes" then
				zz=replace(zzs,num_inputted,"<font color='#FF0000'>"&num_inputted&"</font>")
				else
    			zz=replace(zzs,sousuo,"<font color='#FF0000'>"&sousuo&"</font>")
				end if
				%>
                <td><table width="579" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="559"><a href="#" class="STYLE2" onClick="javascript:window.open('introduce10.asp?id=<%=rs("id")%>','')"><%=rs("name1")%> </a></td>
                    </tr>
                    <tr>
                      <td class="STYLE2"><%if len(zz)>150 then %>
                        <%=left(zz,150)&"....."%>
                        <%else%>
                        <%=zz%>
                        <%end if %>                      </td>
                    </tr>
                    <tr>
                      <td class="STYLE2"><%=rs("beizhu")%> </td>
                    </tr>
                    <tr>
                      <td class="STYLE2"><%=rs("type1")%> </td>
                    </tr>
                    <tr>
                      <td height="18"></td>
                    </tr>
                    <%
			  rs.movenext
			  if rs.eof then exit for
			  next
			  %>
                    <tr>
                      <td height="25" class="STYLE2"><div align="right" class="STYLE2">[<%=page%>/<%=rs.pagecount%>]&nbsp;[共<font color="red"><%=rs_total%></font>条记录]&nbsp;

                          <% if page<>1 then
				if num_inputted = "Yes" then%>
                          <a href=<%=path%>?page=1&num_inputted=<%=num_inputted%>>第一页</a> <a href=<%=path%>?page=<%=(page-1)%>&num_inputted=<%=num_inputted%>> 上一页</a>
						<%else%>
                          <a href=<%=path%>?page=1&sousuo=<%=sousuo%>>第一页</a> <a href=<%=path%>?page=<%=(page-1)%>&sousuo=<%=sousuo%>> 上一页</a>
                          <%end if
						  end if %>
                          <%if page<>rs.pagecount then
						  if num_inputted = "Yes" then%>
                          <a href=<%=path%>?page=<%=(page+1)%>&num_inputted=<%=num_inputted%>>下一页</a> <a href=<%=path%>?page=<%=rs.pagecount%>&num_inputted=<%=num_inputted%>>最后一页</a>
						<%else%>
                          <a href=<%=path%>?page=<%=(page+1)%>&sousuo=<%=sousuo%>>下一页</a> <a href=<%=path%>?page=<%=rs.pagecount%>&sousuo=<%=sousuo%>>最后一页</a>
                          <%end if
						  end if%>
                      </div></td></tr>
                  </table></td>
              </tr>
            </table>
            <p>&nbsp;</p>
          </td>
        </tr>
      </table></td>
  </tr>
</table>
<%
end if
end if
end if
%>
</body>
</html>
