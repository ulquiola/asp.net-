<meta http-equiv="Content-Type" content="text/html; charSet=gb2312">
<% 
    riqi=cint(Request.QueryString("date"))
	If Session("year")="" Then '�û�δʹ������ѡ������,��Ĭ��Ϊϵͳ����
		Session("year")=cint(year(date()))
		Session("month")=cint(month(date()))
		Session("day")=cint(day(date()))
	End If
	If Session("add")<>"" Then '�鿴��һ�µ�����
		Session("add")=""
		Session("month")=cint(Session("month"))+1
		If Session("month")=13 Then 
			Session("month")=1 
			Session("year")=cint(Session("year"))+1
		End If 	
	End If
	If Session("minus")<>"" Then  '�鿴��һ�µ�����
		Session("minus")=""
		Session("month")=Session("month")-1 
		If Session("month")=0 Then 
			Session("month")=12 
			Session("year")=Session("year")-1
		End If 	
	End If	 
	If Request.Querystring("date")<>"" Then 
		Session("day")=riqi
	End If 	
	months=Session("month")
	If months=1 or months=3 or months=5 or months=7 or months=8 or months=10 or months=12 Then 
		sum=31
	Else
		If months=2 Then
			If years Mod 100=0 and years Mod 4=0 or years Mod 400=0 Then  '�ж����겢ȷ�����·ֵ�����
				sum=29
			Else
				sum=28
			End If 
		Else 
			sum=30
		End If 
	End If 
	if Session("day")>sum then 
		Session("day")=sum 
		riqi=sum 
	end if 
	times=cdate(Session("year")&"-"&Session("month")&"-"&Session("day"))
	If Request.QueryString("date")<>"" Then Session("times")=times
%>
<table width="200" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF"> 
  <tr> 
    <td><table width="200" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="27"><img src="images/top1.gif" width="27" height="21"></td>
        <td align="center" background="images/top2.gif">&nbsp;</td>
        <td width="37"><img src="images/top3.gif" width="37" height="21"></td>
      </tr>
    </table></td> 
  </tr> 
  <tr> 
    <td>
	 <table width="200" border="0" cellpadding="0" cellspacing="0"> 
        <tr> 
          <td align="center">
		   <table width="200" border="0" align="center" cellpadding="0" cellspacing="0"
		    bordercolor="#BDBAAD"> 
              <tr> 
                <td width="20"><div align="center">��</div></td> 
                <td width="20"><div align="center">һ</div></td> 
                <td width="20"><div align="center">��</div></td> 
                <td width="20"><div align="center">��</div></td> 
                <td width="20"><div align="center">��</div></td> 
                <td width="20"><div align="center">��</div></td> 
                <td width="20"><div align="center">��</div></td> 
              </tr> 
              <tr> 
                <%
				years=Year(times)
				months=Month(times)
				If months=1 or months=3 or months=5 or months=7 or months=8 or months=10 or months=12 Then 
					sum=31
				Else
					If months=2 Then
						If years Mod 100<>0 and years Mod 4=0 or years Mod 400=0 Then  '�ж��Ƿ�Ϊ����
							sum=29
						Else
							sum=28
						End If 
					Else 
						sum=30
					End If 
				End If 
				'���㵱�ն�Ӧ������ 
				today=WeekDay(times) '
				week=today-1 
				'���㵱���Ǽ���
				days=Day(times)
				'���㵱��1�������ڼ�
				If days-today<0 Then 
					mi=today-days+1
				Else 
					mi=7-((days-today) mod 7)+1
				End If 
				If mi=0 Then mi=7 End If
				'���㵱�¹�ռ���������� 
				mx=cint((sum+mi-1)/7)
				a=(sum+mi-1) mod 7
				If a<>0 and (cdbl(a)/7)<0.5 Then mx=mx+1 End If
				If mi<>1 and mi<>8 Then
					For i=1 To mi-1 
			%> 
                <td> <table border="0" cellspacing="0" cellpadding="0"> 
                    <tr> 
                      <td width="68" height="10"></td> 
                    </tr> 
                    <tr> 
                      <td height="10">&nbsp;</td> 
                    </tr> 
                  </table></td> 
                <% 
					Next
				End If 
				sum1=8-mi
				For m=1 To sum1 
			    %> 
                <td> <table border="0" align="center" cellpadding="0" cellspacing="0"> 
                    <tr> 
                      <td width="68" height="10" bgcolor="<% If riqi=0 Then %>
					  <% If m=days Then%><%="#ffcc99"%><% End If %> <%Else%>
					  <% If m=riqi Then %><%="#ffcc99"%> <% End If %> <% End If %>">
					  <div align="center"><a href="web_blog_list.asp?date=<%=m%>"><%=m%></a></div></td> 
                    </tr> 
                    <tr> 
                      <td height="7"></td> 
                    </tr> 
                  </table></td> 
                <%Next%> 
              </tr> 
              <!-- �ڶ����������ڶ��� --> 
              <% 
				If mx=6 Then '������Ϊ6����ʾ
	 				For j=2 To mx-1 
			  %> 
              <tr> 
						<% For i=1 To 7 %>
                <td> <table border="0" align="center" cellpadding="0" cellspacing="0"> 
                    <tr> 
                      <td width="68" height="10" bgcolor="<% If riqi=0 Then %>
					  <% If (j-2)*7+(8-mi)+i=days Then%><%="#ffcc99"%> <% End If %>
					   <%Else%><% If (j-2)*7+(8-mi)+i=riqi Then %><%="#ffcc99"%>
					    <% End If %> <% End If %>">
					   <div align="center"> 
                          <%If (j-2)*7+(8-mi)+i<=sum Then %> 
                          <a href="web_blog_list.asp?date=<%=(j-2)*7+(8-mi)+i%>">
						   <%=(j-2)*7+(8-mi)+i%></a> 
                          <% End If %> 
                        </div></td> 
                    </tr> 
                    <tr> 
                      <td height="7"></td> 
                    </tr> 
                  </table></td> 
							<%Next%> 
              </tr> 
               <% 
						Next 
					Else 
						For j=2 To mx 
				%> 
              <tr> 
							<%For i=1 To 7 %>
                <td> <table border="0" align="center" cellpadding="0" cellspacing="0"> 
                    <tr> 
                      <td width="68" height="10" bgcolor="<% If riqi=0 Then %>
					  <% If (j-2)*7+(8-mi)+i=days Then%> <%="#ffcc99"%><% End If %>
					   <%Else%>
					   <% If (j-2)*7+(8-mi)+i=riqi Then %>
					   <%="#ffcc99"%> <% End If %> <% End If %>">
					   <div align="center"> 
                          <%If (j-2)*7+(8-mi)+i<=sum Then %> 
                          <a href="web_blog_list.asp?date=<%=(j-2)*7+(8-mi)+i%>">
						  <%=(j-2)*7+(8-mi)+i%></a> 
                          <% End If %> 
                        </div></td> 
                    </tr> 
                    <tr> 
                      <td height="7"></td> 
                    </tr> 
                  </table></td> 
								<%Next%> 
              </tr> 
							<%Next%> 
              <tr> 
                <% 
						End If 
					'���������һ����ʾ
					If mx=6 Then 
						For i=1 To sum-((mx-2)*7+(8-mi)) 
				%> 
                <td> <table border="0" align="center" cellpadding="0" cellspacing="0"> 
                    <tr> 
                      <td height="10" bgcolor="<% If riqi=0 Then %>
					  <% If (mx-2)*7+(8-mi)+i=days Then%><%="#ffcc99"%> <% End If %>
					  <%Else%><% If (mx-2)*7+(8-mi)+i=riqi Then %>
					   <%="#ffcc99"%> <% End If %> <% End If %>">
					    <a href="web_blog_list.asp?date=<%=(mx-2)*7+(8-mi)+i%>">
						<%=(mx-2)*7+(8-mi)+i%></a> </td> 
                    </tr> 
                    <tr> 
                      <td height="2"></td> 
                    </tr> 
                  </table></td> 
						<%Next%> 
                <td colspan="<%=7-(sum-((mx-2)*7+(8-mi))) %>">
				 <table border="0" align="center" cellpadding="0" cellspacing="0"  width=100%> 
                    <tr> 
                      <td  height="10" align="right">
					   <img src="Images/Forum_up.gIf" width="8" height="7"
					    onClick="javascript:minus.submit()">
						 <%=Session("year")%>��<%=Session("month")%>��
					  <img src="Images/Forum_nav.gIf" width="8" height="7"
						   onClick="javascript:add.submit()"> </td> 
                    </tr> 
                    <tr> 
                      <td height="7" background="images/bottom1.gif"></td> 
                    </tr> 
                  </table></td> 
                <% 
					Else 
				%> 
                <td colspan="7">
				<table border="0" align="center" cellpadding="0" cellspacing="0" width=100%> 
                    <tr> 
                      <td height="15" align="right">
					   <img src="Images/Forum_up.gIf" width="8" height="7"
					    onClick="javascript:minus.submit()">
						 <%=Session("year")%>��<%=Session("month")%>��
					  <img src="Images/Forum_nav.gIf" width="8" height="7"
						   onClick="javascript:add.submit()"> </td> 
                    </tr> 
                    <tr> 
                      <td height="7"></td> 
                    </tr> 
                  </table></td> 
					<%End If%> 
                <Form name="add" method="post" action="web_blog_list.asp?date=<%=day(times)%>"> 
                  <input type="hidden" name="add" value="1"> 
                </Form> 
                <Form name="minus" method="post" action="web_blog_list.asp"> 
                  <input type="hidden" name="minus" value="1"> 
                </Form> 
              </tr>	
			  <tr><td>  </td></tr>		    
            </table></td> 
        </tr> 
      </table></td> 
  </tr> 
</table> 
<table width="220" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/bottom1.gif">&nbsp;</td>
  </tr>
</table>
