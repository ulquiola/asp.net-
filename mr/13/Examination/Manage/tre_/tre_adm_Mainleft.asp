<%
getproid = replace(trim(request("profession")),"'","''")
if(getproid <> "")then
session("proid") = getproid
response.write("<script>location.href = '../adm_Mainleft.asp'</script>")
end if
%>
