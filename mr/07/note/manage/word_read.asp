<link rel="stylesheet" href="../css/css.css" />
<table width="950" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td colspan="3"><!--#include file="m_header.html"--></td>
  </tr>
  <tr>
    <td width="190" valign="top" bgcolor="#f0f0f0"><!--#include file="m_left.asp"--></td>
    <td width="757" align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td height="59" background="../images/readword.jpg">&nbsp;</td>
      </tr>
    </table>
      <table width="700"  border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td width="613" height="223" align="left" valign="top">&nbsp;
		<%
		Set FSObject=Server.CreateObject("Scripting.FileSystemObject")
		Set FileObject=FSObject.GetFile(Server.MapPath("../filterwords.txt"))
		Set TF=FileObject.OpenAsTextStream(1)
		Response.Write "Ãô¸Ð´ÊÈçÏÂ£º "
		Do while not TF.AtEndOfStream
			Response.Write TF.Readline
			Response.Write "¡¢"
		Loop
		TF.close
		Set TF=Nothing
		Set FSObject=Nothing
		%>		
		</td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td colspan="3"><!--#include file="footer.html"--></td>
  </tr>
</table>
