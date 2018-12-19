<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<table border="0" cellpadding="0" cellspacing="0" width="590">
      <tr>
        <td colspan="3"><img name="index_7_r1_c1" src="images/index_7_r1_c1.jpg" width="590" height="49" border="0" alt=""></td>
      </tr>
      <tr>
        <td colspan="3"><img name="index_7_r2_c1" src="images/index_7_r2_c1.jpg" width="590" height="9" border="0" alt=""></td>
      </tr>
      <tr>
        <td width="16" background="images/index_7_r3_c1.jpg">&nbsp;</td>
        <td width="564"><table width="96%" height="153"  border="0" align="center" cellpadding="0" cellspacing="0">
		  <tr>
<%
i=0
sql="select top 2 * from [shangpin] order by id desc"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
do while not rs.eof
	i=i+1
%>
            <td height="89"><table width="255"  border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="130" rowspan="6"><div align="center"><a href="lookpro.asp?id=<%=rs("id")%>" target="_blank"><img src="upfile/<%=rs("tupian")%>" width="110" height="130" border="0"></a></div></td>
                <td width="20" height="16">&nbsp;</td>
                <td width="113"><font color="EF9C3E">【<%=rs("mingcheng")%>】</font></td>
              </tr>
              <tr>
                <td height="16">&nbsp;</td>
                <td><font color="910800">【市场价：<%=rs("shichang")%>】</font></td>
              </tr>
              <tr>
                <td height="16">&nbsp;</td>
                <td><font color="DD4679">【会员价：<%=rs("huiyuan")%>】</font></td>
              </tr>
              <tr>
                <td height="16">&nbsp;</td>
                <td><a href="lookpro.asp?id=<%=rs("id")%>" target="_blank">【查看信息</a>】</td>
              </tr>
              <tr>
                <td height="16">&nbsp;</td>
                    <td>【<a href="gouwu.asp?ProdId=<%=rs("id")%>">购买商品</a>】</td>
              </tr>
              <tr>
                <td height="16">&nbsp;</td>
                <td><font color="13589B">【浏览次数：<%=rs("cishu")%>】</font></td>
              </tr>
            </table></td>
<%if i mod 2=0 then%>
          </tr>
<%
end if
rs.movenext
loop
rs.close
set rs=nothing
%>           
        </table></td>
        <td width="10" background="images/index_7_r3_c3.jpg">&nbsp;</td>
      </tr>
      <tr>
        <td colspan="3"><img name="index_7_r4_c1" src="images/index_7_r4_c1.jpg" width="590" height="7" border="0" alt=""></td>
      </tr>
      <tr>
        <td colspan="3"><img name="index_7_r5_c1" src="images/index_7_r5_c1.jpg" width="590" height="49" border="0" alt=""></td>
      </tr>
      <tr>
        <td colspan="3"><img name="index_7_r6_c1" src="images/index_7_r6_c1.jpg" width="590" height="7" border="0" alt=""></td>
      </tr>
      <tr>
        <td><img name="index_7_r7_c1" src="images/index_7_r7_c1.jpg" width="16" height="168" border="0" alt=""></td>
        <td><table width="96%" height="153"  border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
<%
i=0
sql="select top 2 * from [shangpin] where dengji='1' order by id desc"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
do while not rs.eof
	i=i+1
%>
            <td height="89"><table width="255"  border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="130" rowspan="6"><div align="center"><a href="lookpro.asp?id=<%=rs("id")%>" target="_blank"><img src="upfile/<%=rs("tupian")%>" width="110" height="130" border="0"></a></div></td>
                <td width="20" height="16">&nbsp;</td>
                <td width="113"><font color="EF9C3E">【<%=rs("mingcheng")%>】</font></td>
              </tr>
              <tr>
                <td height="16">&nbsp;</td>
                <td><font color="910800">【市场价：<%=rs("shichang")%>】</font></td>
              </tr>
              <tr>
                <td height="16">&nbsp;</td>
                <td><font color="DD4679">【会员价：<%=rs("huiyuan")%>】</font></td>
              </tr>
              <tr>
                <td height="16">&nbsp;</td>
                <td><a href="lookpro.asp?id=<%=rs("id")%>" target="_blank">【查看信息</a>】</td>
              </tr>
              <tr>
                <td height="16">&nbsp;</td>
                    <td>【<a href="gouwu.asp?ProdId=<%=rs("id")%>">购买商品</a>】</td>
              </tr>
              <tr>
                <td height="16">&nbsp;</td>
                <td><font color="13589B">【浏览次数：<%=rs("cishu")%>】</font></td>
              </tr>
            </table></td>
            <%if i mod 2=0 then%>
          </tr>
<%
end if
rs.movenext
loop
rs.close
set rs=nothing
%>
        </table></td>
        <td><img name="index_7_r7_c3" src="images/index_7_r7_c3.jpg" width="10" height="168" border="0" alt=""></td>
      </tr>
      <tr>
        <td colspan="3"><img name="index_7_r8_c1" src="images/index_7_r8_c1.jpg" width="590" height="7" border="0" alt=""></td>
      </tr>
      <tr>
        <td colspan="3"><img name="index_7_r9_c1" src="images/index_7_r9_c1.jpg" width="590" height="48" border="0" alt=""></td>
      </tr>
      <tr>
        <td colspan="3"><img name="index_7_r10_c1" src="images/index_7_r10_c1.jpg" width="590" height="10" border="0" alt=""></td>
      </tr>
      <tr>
        <td><img name="index_7_r11_c1" src="images/index_7_r11_c1.jpg" width="16" height="168" border="0" alt=""></td>
        <td><table width="96%" height="153"  border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
<%
i=0
sql="select top 2 * from [shangpin] order by cishu desc"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
do while not rs.eof
	i=i+1
%>
            <td height="89"><table width="255"  border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="130" rowspan="6"><div align="center"><a href="lookpro.asp?id=<%=rs("id")%>" target="_blank"><img src="upfile/<%=rs("tupian")%>" width="110" height="130" border="0"></a></div></td>
                  <td width="20" height="16">&nbsp;</td>
                  <td width="113"><font color="EF9C3E">【<%=rs("mingcheng")%>】</font></td>
                </tr>
                <tr>
                  <td height="16">&nbsp;</td>
                  <td><font color="910800">【市场价：<%=rs("shichang")%>】</font></td>
                </tr>
                <tr>
                  <td height="16">&nbsp;</td>
                  <td><font color="DD4679">【会员价：<%=rs("huiyuan")%>】</font></td>
                </tr>
                <tr>
                  <td height="16">&nbsp;</td>
                  <td><a href="lookpro.asp?id=<%=rs("id")%>" target="_blank">【查看信息</a>】</td>
                </tr>
                <tr>
                  <td height="16">&nbsp;</td>
                    <td>【<a href="gouwu.asp?ProdId=<%=rs("id")%>">购买商品</a>】</td>
                </tr>
                <tr>
                  <td height="16">&nbsp;</td>
                  <td><font color="13589B">【浏览次数：<%=rs("cishu")%>】</font></td>
                </tr>
            </table>
</td>
            <%if i mod 2=0 then%>
          </tr>
<%
end if
rs.movenext
loop
rs.close
set rs=nothing
%>
        </table></td>
        <td><img name="index_7_r11_c3" src="images/index_7_r11_c3.jpg" width="10" height="168" border="0" alt=""></td>
      </tr>
      <tr>
        <td colspan="3"><img name="index_7_r12_c1" src="images/index_7_r12_c1.jpg" width="590" height="7" border="0" alt=""></td>
      </tr>
      <tr>
        <td height="58" colspan="3">&nbsp;</td>
      </tr>
    </table>