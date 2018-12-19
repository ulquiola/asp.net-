<%
function Display(rs_qyNO)
set conn=server.CreateObject("ADODB.Connection")
conn.open("Driver={Microsoft Access Driver (*.mdb)};PWD=111;DBQ=" & server.MapPath("../Customer.mdb"))
%>
<%
sql_qyNO="select* from Q_area where ID="& qyNO
set rs_qyNO=conn.execute(sql_qyNO)
if rs_qyNO("father")<>0 then
	sql_father_1="select* from Q_area where ID="& rs_qyNO("father")
	set rs_father_1=conn.execute(sql_father_1) 
	if rs_father_1("father")<>0 then
	sql_father_1_1="select* from Q_area where ID="& rs_father_1("father")
	set rs_father_1_1=conn.execute(sql_father_1_1)
	else 
	sql_father_1_1="select* from Q_area where ID=0"
	set rs_father_1_1=conn.execute(sql_father_1_1)
	end if
	else
	sql_father_1="select* from Q_area where ID=0"
	set rs_father_1=conn.execute(sql_father_1)
	sql_father_1_1="select* from Q_area where ID=0"
	set rs_father_1_1=conn.execute(sql_father_1_1)
	end if
	 %>
	<% if not rs_father_1_1.eof then %>
	<select name="town_F" id="select1">
        <option value="<%=rs_father_1_1("Name")%>"><%=rs_father_1_1("Name")%></option></select>
	(<%=rs_father_1_1("TypeName")%>)
	<% end if %>
	<% if not rs_father_1.eof then %>
	<select name="city_F" id="select2">
        <option value="<%=rs_father_1("Name")%>"><%=rs_father_1("Name")%></option></select>
	(<%=rs_father_1("TypeName")%>) 
	<% end if %>
	<select name="province_F" id="select3">
         <option value="<%=(rs_qyNO("Name"))%>"><%=(rs_qyNO("Name"))%></option></select>
	(<%=(rs_qyNO("TypeName"))%>)
<% End function %>
