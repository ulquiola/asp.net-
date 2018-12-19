<html><head><title>附件上传</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
body {	background-color: #E8F1FF;}
body,td,th {
	font-size: 12px;
}
-->
</style>
<script language="javascript">
<!--
function mysub(){
  esave.style.visibility="visible";
}
-->
</script>
</head>
<body>
<%
uppath=request.QueryString("uppath")&"/"			
filelx=request.QueryString("filelx")				
formName=request.QueryString("formName")			
EditName=request.QueryString("EditName")			
%>
<form name="form1" method="post" action="upfile_deal.asp" enctype="multipart/form-data" >
<div id="esave" style="position:absolute; top:18px; left:40px; z-index:10; visibility:hidden"> 
<TABLE WIDTH=340 BORDER=0 CELLSPACING=0 CELLPADDING=0>
<TR><td width=20%></td>
<TD bgcolor=#ff0000 width="60%"> 
<TABLE WIDTH=100% height=120 BORDER=0 CELLSPACING=1 CELLPADDING=0>
<TR> 
<td bgcolor=#ffffff align=center><font color=red class="fo">上传文件...</font></td>
</tr>
</table>
</td><td width=20%></td>
</tr></table></div>
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
<tr bgcolor="#E8F1FF"> 
<td align="center" id="upid" height="80">
        <input type="hidden" name="filepath" value="<%=uppath%>">
        <input type="hidden" name="filelx" value="<%=filelx%>">
        <input type="hidden" name="EditName" value="<%=EditName%>">
        <input type="hidden" name="FormName" value="<%=formName%>">
        <input type="hidden" name="act" value="uploadfile">
        <span class="fo">选择文件:</span>&nbsp;
  <input type="file" name="file1" size="40" value="">&nbsp;<input type="submit" name="Submit" value="开始上传" onclick="javascript:mysub()"></td>
</tr>
</table>
</form>
</body>
</html>