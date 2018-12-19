<%
	curDay = int(day(now()))
	if curDay mod 2 = 0 then
		response.redirect("double.asp")
	else
		response.redirect("sinple.asp")
	end if
%>