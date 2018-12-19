<link rel="stylesheet" href="../css/css.css" />
<script language="JavaScript">
function check(myform){
	if(myform.txt_word.value==""){
		alert("敏感词不能为空");myform.txt_word.focus();return false;
	}
}
</script>
<table width="950" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td colspan="3"><!--#include file="m_header.html"--></td>
  </tr>
  <tr>
    <td width="190" valign="top" bgcolor="#f0f0f0"><!--#include file="m_left.asp"--></td>
    <td width="757" align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" height="59" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td background="../images/addword.jpg">&nbsp;</td>
      </tr>
    </table>
    <table width="700"  border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td width="613" height="223" align="center"><br />
            <table width="700" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td align="center"><form action="" method="post"  name="myform">
                    <table width="560" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td valign="middle" align="right" width="20%">添加敏感词：<br /></td>
                        <td width="80%"><input name="txt_word" type="text" id="txt_word" size="50" />
                          &nbsp;
                        <input name="btn_tj" type="submit" class="btn1" id="btn_tj" onclick="return check(myform);" value="添加敏感词" /></td>
                      </tr>
                    </table>
                </form></td>
              </tr>
          </table></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td colspan="3"><!--#include file="footer.html"--></td>
  </tr>
</table>
<%
if(request.Form("btn_tj")<>"")then
	Set FSObject=Server.CreateObject("Scripting.FileSystemObject")
	Set TF=FSObject.OpenTextFile(Server.MapPath("../filterwords.txt"),8,true)
	txt_word=request.Form("txt_word")
	TF.Writeline(txt_word)
	Response.Write "<br>"
	TF.close
	Set TF=Nothing
	Set FSObject=Nothing
%>
<script language="javascript">alert("敏感词汇添加成功！");window.location.href='word_add.asp';</script>
<%	
end if
%>

