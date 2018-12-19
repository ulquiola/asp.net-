<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="inc/Validate.asp"-->
<!--#include file="Conn/conn.asp"-->
<%
function level(num)
	select case num
		case "1"
			level = "一级评语"
		case "2"
			level = "二级评语"
		case "3"
			level = "三级评语"
		case "4"
			level = "四级评语"
		case "5"
			level = "五级评语"
		case "6"
			level = "六级评语"
		case "7"
			level = "七级评语"
		case "8"
			level = "八级评语"
	end select
end function
getkey=Session("StuID")
set rs = server.createobject("adodb.recordset") 
rssql = "Select * from tb_StuResult where 考号='"&getkey&"'"
rs.open rssql,conn,1,3
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>查询成绩</title>
<link rel="stylesheet" type="text/css" href="Css/style.css">
<style type="text/css">
<!--
.STYLE1 {
	color: #333333;
	font-weight: bold;
}
body {
	margin-top: 0px;
}
-->
</style>
</head>
<body leftmargin="0">
<table cellpadding="0" cellspacing="0" border="0" width="100%" height="90%">
  <tr>
    <td align="center" height="18"><span class="STYLE1">查询成绩</span> </td>
  </tr>
  <tr>
    <td height="354" valign="top"><table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
        <tr>
          <td valign="top">
            <table width="100%"  border="1" cellspacing="0" cellpadding="0" align="center"
 bordercolor="#FFFFFF" bordercolordark="#1985DB" bordercolorlight="#FFFFFF">
              <tr>
                <td height="26" colspan="6"  align="center"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="26%" height="24" align="center" bgcolor="#00CCFF"><span style="color:black;"><strong>考生ID号</strong></span></td>
                    <td width="40%" align="center">
						<strong><span style="color:black;">
						<%
							if not rs.eof then
						 		response.Write(rs("考生ID号"))
							else
								response.Write("<font color=red>没有符合条件的记录！</font>")
								response.End()
							end if
						%></span></strong></td>
                    <td width="34%">&nbsp;</td>
                  </tr>
                </table></td>
              </tr>
              <tr>
                <td width="26%" height="24"  align="center" nowrap bgcolor="#00CCFF" style="color:black;"><strong>考号</strong></td>
                <td width="11%" height="26"  align="center" nowrap bgcolor="#00CCFF" style="color:black;">级别</td>
                <td width="19%" height="26"  align="center" nowrap bgcolor="#00CCFF" style="color:black;">姓名</td>
                <td width="16%" height="26"  align="center" nowrap bgcolor="#00CCFF" style="color:black;">性别</td>
                <td width="14%" height="26"  align="center" nowrap bgcolor="#00CCFF" style="color:black;">年级</td>
                <td width="14%" height="26"  align="center" nowrap bgcolor="#00CCFF" style="color:black;">考点代码</td>
              </tr>
              <tr>
                <td height="24"  align=center nowrap><%=rs("考号")%></td>
                <td height="24"  align=center nowrap style="color:black;"><%=rs("级别")%>&nbsp;</td>
                <td height="24"  align=center nowrap style="color:black;"><%=rs("姓名")%>&nbsp;</td>
                <td height="24"  align=center nowrap style="color:black;"><%=rs("性别")%>&nbsp;</td>
                <td height="24"  align=center nowrap style="color:black;"><%=rs("年级")%>&nbsp;</td>
                <td height="24"  align=center nowrap style="color:black;"><%=rs("考点代码")%>&nbsp;</td>
              </tr>
            </table>
            <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#1985DB">
              <tr>
                <td width="7%" height="24" align="center" bgcolor="#00CCFF">科目</td>
                <td width="19%" align="center" bgcolor="#00CCFF">得分</td>
                <td width="11%" align="center" bgcolor="#00CCFF">等级</td>
                <td width="63%" align="center" bgcolor="#00CCFF">评语</td>
              </tr>
             
                <%
				'多表串联
				set rs1 = server.createobject("adodb.recordset") 
				rssql = "SELECT tb_StuResult.*, tb_ear."&level(rs("级别"))&" as 听力评语, tb_talk."&level(rs("级别"))&" as 口语评语, tb_write."&level(rs("级别"))&" as 作文评语, tb_zonghe."&level(rs("级别"))&" as 综合评语 "&_
				"FROM tb_zonghe INNER JOIN ("&_
					"tb_write INNER JOIN ("&_
						"tb_talk INNER JOIN ("&_
							"tb_ear INNER JOIN tb_StuResult ON tb_ear.听力等级 = tb_StuResult.听力等级"&_
						") ON tb_talk.口语等级 = tb_StuResult.口语等级"&_
					") ON tb_write.作文等级 = tb_StuResult.作文等级"&_
				") ON tb_zonghe.综合等级 = tb_StuResult.综合等级"&_
				" where tb_sturesult.考号 = '130008100002'"
				rs1.open rssql,conn,1,3
				%>
			<tr>
                <td align="center" bgcolor="#00CCFF">口语</td>
                <td align="center"><%=rs1("口语得分")%>&nbsp;</td>
                <td align="center"><%=server.htmlencode(rs1("口语等级"))%>&nbsp;</td>
                <td><%=rs1("口语评语")%>&nbsp;</td>
              </tr>
              <tr>
                <td align="center" bgcolor="#00CCFF">听力</td>
                <td align="center"><%=server.htmlencode(rs1("听力得分"))%>&nbsp;</td>
                <td align="center"><%=server.htmlencode(rs1("听力等级"))%>&nbsp;</td>
                <td><%=rs1("听力评语")%></td>
              </tr>
                <td align="center" bgcolor="#00CCFF">综合</td>
                <td align="center" bgcolor="#CCCCCC"><table width="100%" border="0" cellspacing="1" cellpadding="0">
                    <tr>
                      <td width="49%" align="center" bgcolor="#FFFFFF"><span style="color:black;">综合题Ⅰ</span></td>
                      <td width="22%" align="center" bgcolor="#FFFFFF"><span style="color:black;"><%=rs1("综合题Ⅰ")%></span>&nbsp;</td>
                      <td width="29%" rowspan="5" align="center" bgcolor="#FFFFFF"><%=server.htmlencode(rs1("综合得分"))%>&nbsp;</td>
                    </tr>
                    <tr>
                      <td align="center" bgcolor="#FFFFFF"><span style="color:black;">综合题Ⅱ</span></td>
                      <td align="center" bgcolor="#FFFFFF"><span style="color:black;"><%=rs1("综合题Ⅱ")%></span>&nbsp;</td>
                    </tr>
                    <tr>
                      <td align="center" bgcolor="#FFFFFF"><span style="color:black;">综合题Ⅲ</span></td>
                      <td align="center" bgcolor="#FFFFFF"><span style="color:black;"><%=rs1("综合题Ⅲ")%></span>&nbsp;</td>
                    </tr>
                    <tr>
                      <td align="center" bgcolor="#FFFFFF"><span style="color:black;">综合题Ⅳ</span></td>
                      <td align="center" bgcolor="#FFFFFF"><span style="color:black;"><%=rs1("综合题Ⅳ")%></span>&nbsp;</td>
                    </tr>
                    <tr>
                      <td align="center" bgcolor="#FFFFFF"><span style="color:black;">综合题Ⅴ</span></td>
                      <td align="center" bgcolor="#FFFFFF"><span style="color:black;"><%=rs1("综合题Ⅴ")%></span>&nbsp;</td>
                    </tr>
                  </table></td>
                <td align="center"><%=server.htmlencode(rs1("综合等级"))%>&nbsp;</td>
                <td><%=rs1("综合评语")%>
                &nbsp;</td>
              </tr>
              <tr>
                <td align="center" bgcolor="#00CCFF">作文</td>
                <td align="center"><%=rs1("作文得分")%>&nbsp;</td>
                <td align="center"><%=server.htmlencode(rs1("作文等级"))%>&nbsp;</td>
                <td><%=rs1("作文评语")%>
                &nbsp;</td>
              </tr>
              <tr>
                <td height="26" align="center" bgcolor="#00CCFF"><strong>总成绩</strong></td>
                <td align="center"><strong><%=server.htmlencode(rs1("总分"))%></strong>&nbsp;</td>
                <td align="center"><strong><%=server.htmlencode(rs1("总评"))%></strong>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
          </table></td>
        </tr>
    </table></td>
  </tr>
</table>
</body>
</html>
