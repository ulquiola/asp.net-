<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<!--#include file="../conn/conn.asp"-->
<%
if request.Form("name1")<>"" then
name1=request.Form("name1")
department=request.Form("department")
enroltype=request.Form("enroltype")
enrolremark=request.Form("enrolremark")
ss=now()
'获取登记类型
select case enroltype
       case "上班登记"
	   '应用cdate函数将字符串转换为日期数据类型
	   ee=cdate(date()&" 08:20:00")
	   if ee-ss>=0 then
	   	state1="正常"
	   else
	   	state1="迟到"
	   end if
		sql="insert into Tab_onduty(name1,department,enroltype,enrolremark,definetime,enroltime,state) values('"&name1&"','"&department&"','"&enroltype&"','"&enrolremark&"','"&ee&"','"&ss&"','"&state1&"')"
		conn.execute(sql)
		case "下班登记"
		'应用cdate函数将字符串转换为日期数据类型
		ee=cdate(date()&" 17:10:00")
		ee1=cdate(date()&" 20:30:00")
	   if ee-ss<=0 and ee1-ss<=0 then
	   state1="加班"
	   else if ee-ss<=0 and ee1-ss>0 then
	   	state1="正常"
	   else
	   	state1="早退"
	    end if
		end if	   
			sql="insert into Tab_onduty(name1,department,enroltype,enrolremark,definetime,enroltime,state) values('"&name1&"','"&department&"','"&enroltype&"','"&enrolremark&"','"&ee&"','"&ss&"','"&state1&"')"
		conn.execute(sql)
end select
%>
<script language="javascript">
alert("登记信息添加成功！");
opener.location.reload();
window.close();
</script>
<%end if%>