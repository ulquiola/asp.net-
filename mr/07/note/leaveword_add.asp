<meta http-equiv="Content-Type" content="text/html;	charset=gb2312">
<link href="css/css.css" rel="stylesheet" type="text/css">
<%
Set FSObject=Server.CreateObject("Scripting.FileSystemObject")
Set FileObject=FSObject.GetFile(Server.MapPath("filterwords.txt"))
Set TF=FileObject.OpenAsTextStream(1)			'以只读方式打开文本文件
Do while not TF.AtEndOfStream					'循环输出敏感词汇
	word_str=word_str&TF.Readline&"、"
Loop
word_str=left(word_str,len(word_str)-1)
TF.close										'关闭文件
Set TF=Nothing									'清空文件资源
Set FSObject=Nothing							'清空系统资源
%>		
<script language="JavaScript">
function check_form(form1){
		if(form1.user_name.value==""){
			alert("昵称不能为空！");form1.user_name.focus();return false;
		}
		if(form1.title.value==""){
			alert("标题不能为空!");form1.title.focus();return false;
		}
		if(form1.content.value=="")	{
			alert("内容不能为空!");form1.content.focus();return false;
		}
		var str1=form1.hiddenField.value;				//接收隐藏域的值
		var content=form1.content.value;				//接收留言信息
		arr=str1.split("、");							//以“、”拆分存储在文本文件中的敏感词
		for(i=0;i<arr.length;i++){						//应用For循环语句对敏感词进行判断
			str1=content.indexOf(arr[i],0);				//判断传递的留言信息中是否含有敏感词
			if(str1>=0){								//如果有查询结果，则弹出提示信息
				alert("留言信息中有敏感词汇!");form1.content.focus();return false;
			}
		}
	form1.submit();										//提交表单信息
}
</script>

<table width="950" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><!--#include file="header.html"--></td>
  </tr>
  <tr>
    <td valign="top"><table width="950" align="center"cellpadding="0" cellspacing="0" style="border:0;align:center;">
        <tr>
          <td width="949" align="right"  bgcolor="#FFFFFF"><table width="950" border="1" align="center" cellpadding="1" cellspacing="1" bordercolor="#FFFFFF" bgcolor="#FFFFFF">
              <tr>
                <td><table width="100%" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#FFFFFF" bgcolor="#FFFFFF">
              <tr>
                <td height="44" background="images/a_11.jpg">&nbsp;</td>
              </tr></table>
                  <table width="761" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#FEFEFE" bgcolor="#FFFFFF">
                    <form action="note_check.asp"  method="post" name="form1" id="form1">
                      <tr>
						<td width="761" height="40" bgcolor="#FFFFFF" style="BACKGROUND-IMAGE: url(image/bar_pink.gif); COLOR: #000000; BACKGROUND-COLOR: #FFFFFF">
                        	<hr color="#8AA925" size="2" style="width:750px;">
						</td>
                      </tr>
                      <tr>
                        <td align="center" bgcolor="#F9F8EF"><table width="749" border="0" align="center" cellpadding="0"  cellspacing="0"  style="BORDER-COLLAPSE: collapse">
                            <tr>
                              <td width="749" height="57" background="images/a_03.jpg">&nbsp;&nbsp;</td>
                            </tr>
                            <tr>
                              <td height="36" colspan="3" align="left" background="images/a_05.jpg" bgcolor="#F9F8EF" scope="col">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;昵&nbsp;&nbsp;称：
                                <input  name="user_name" id="user_name" value=" 匿名" maxlength="64" type="text" />
                                  <span style="COLOR: #ff0000">*</span></td>
                            </tr>
                            <tr>
                              <td height="36" colspan="3" align="left" background="images/a_05.jpg" bgcolor="#F9F8EF">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;标　题：
                                <input maxlength="64" size="30" name="title"  type="text"/>
                                  <span style="COLOR: #ff0000">*</span></td>
                            </tr>
                            <tr>
                              <td height="126" colspan="3" align="left" background="images/a_05.jpg" bgcolor="#F9F8EF">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;内&nbsp;&nbsp;容：
                                <textarea name="content" cols="60" rows="8" id="content" style="background:url(./images/mrbccd.gif)"></textarea>
                                  <span style="COLOR: #ff0000">*</span></td>
                            </tr>
                            <tr>
                              <td height="40" colspan="3" align="left" background="images/a_05.jpg" bgcolor="#F9F8EF">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                <input type="checkbox" name="checkbox" value="1" />
                                给版主的悄悄话
                                <input type="hidden" name="hiddenField"  value="<% response.Write(word_str) %>"/></td>
                            </tr>
                            <tr>
                              <td height="35" background="images/a_07.jpg">&nbsp;&nbsp;</td>
                            </tr>
                        </table></td>
                      </tr>
                      <tr>
                        <td align="center" bgcolor="#F9F8EF"><table width="754" border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td height="58" background="images/a_08.jpg"></td>
                            </tr>
                            <tr>
                              <td height="76" align="center" background="images/a_09.jpg"><table width="534" height="100" border="0"   cellpadding="0" cellspacing="0">
                                  <tr style="VERTICAL-ALIGN: bottom">
                                    <td width="132" 
                style="BORDER-RIGHT: #D9D2B6 1px solid; BORDER-TOP: #D9D2B6 1px solid; BORDER-LEFT: #D9D2B6 1px solid; BORDER-BOTTOM: #D9D2B6 1px solid"><input  type="radio"  checked="checked" value="01" name="head" />
                                        <img  src="images/face/pic/01.gif" width="90" height="90" /></td>
                                    <td width="134" 
                style="BORDER-RIGHT: #D9D2B6 1px solid; BORDER-TOP: #D9D2B6 1px solid; BORDER-LEFT: #D9D2B6 1px solid; BORDER-BOTTOM: #D9D2B6 1px solid"><input  type="radio" value="02" name="head" />
                                        <img   src="images/face/pic/02.gif" width="90" height="90" /></td>
                                    <td width="134" 
                style="BORDER-RIGHT: #D9D2B6 1px solid; BORDER-TOP: #D9D2B6 1px solid; BORDER-LEFT: #D9D2B6 1px solid; BORDER-BOTTOM: #D9D2B6 1px solid"><input  type="radio" value="03" name="head" />
                                        <img  src="images/face/pic/03.gif" width="90" height="90" /></td>
                                    <td width="134"  style="BORDER-RIGHT: #D9D2B6 1px solid; BORDER-TOP: #D9D2B6 1px solid; BORDER-LEFT: #D9D2B6 1px solid; BORDER-BOTTOM: #D9D2B6 1px solid"><input type="radio" value="04" name="head" />
                                     	<img  src="images/face/pic/04.gif" width="90" height="90"></td>
                                  </tr>
                              </table></td>
                            </tr>
                            <tr>
                              <td height="20" background="images/a_10.jpg">&nbsp;</td>
                            </tr>
                          </table>
                            <table width="734" border="0" align="center" cellpadding="0" cellspacing="0">
                              <tr>
                                <td width="703" height="40" align="center"><input name="btn" type="button" class="btn1" id="btn" value="签写留言" onclick="return check_form(form1);"/>
                                  &nbsp;&nbsp;
                                  <input name="reset" type="reset" class="btn1" value="清除留言" /></td>
                              </tr>
                          </table></td>
                      </tr>
                    </form>
                  </table></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td><!--#include file="footer.html"--></td>
  </tr>
</table>
