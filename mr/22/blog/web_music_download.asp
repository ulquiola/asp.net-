<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="include/conn.asp"-->
<%
if request.QueryString("id")<>"" then id=request.QueryString("id")
set rs=server.CreateObject("adodb.recordset")
sqlstr="select Mpath from tab_music where id="&id&""
rs.open sqlstr,conn,1,1
if rs.eof and rs.bof then response.end()
Mpath=rs("Mpath")
rs.close
Set rs=Nothing
%>
<html>
<head>
<title>mp3文件下载</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="css/css.css" />
</head>
<body>
<%call downloadFile(Mpath)  '调用函数downloadFile(Filename),实现下载功能
Function downloadFile(Filename)  
 Filepath =server.MapPath("music")&"\"& Filename
 Response.Buffer = True  
 Response.Clear  
 '创建Stream对象
 Set Objstream = Server.CreateObject("ADODB.Stream")  
 Objstream.Open  
 '数据类型为二进制
 Objstream.Type = 1  
 on error resume next  
 '检查文件是否出错
 Set Objfile = Server.CreateObject("Scripting.FileSystemObject")  
 if not Objfile.FileExists(Filepath) then  %>
  <table width="560" height="395" border="0" cellpadding="0" cellspacing="0" background="../images/file/download.gif">
  <tr>
    <td width="510" valign="top"><table width="100%" height="137"  border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td height="121">&nbsp;</td>
      </tr>
    </table>      
    <table width="83%" height="214"  border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td height="91">
		<% Response.Write("错误提示：<p>" & Filepath & " 不存在！<p>") %> 
		</td>
      </tr>
      <tr>
       <td><div align="center"><br>
		<form name="form1">
        <input name="myclose" type="button" class="Style_button_del" id="myclose"
		 value="关闭窗口" onClick="javascrip:window.close()">
        </form>
		</div></td>
       </tr>
      </table>
    </td>
  </tr>
</table>
 <%
 Response.End() 
 end if  
 '获取文件的大小
 Set myfile = Objfile.GetFile(Filepath)  
 intFilelength = myfile.size  
 Objstream.LoadFromFile(Filepath)  
 if err then  %>
  <table width="560" height="395" border="0" cellpadding="0" cellspacing="0"
   background="../images/file/download.gif">
  <tr>
    <td width="510" valign="top"><table width="100%" height="131"  border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td height="120">&nbsp;</td>
      </tr>
    </table>      
        <table width="83%" height="189"  border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td height="91">
			<% Response.Write("错误提示:<p> " & err.Description & "<p>") %>
 			</td>
          </tr>
          <tr>
            <td height="98"><div align="center">
                <br>
                <input name="myclose" type="button" class="Style_button_del"
				 id="myclose" value="关闭窗口" onClick="javascrip:window.close()">
            </div></td>
          </tr>
      </table>
    </form></td>
  </tr>
</table>
 <%
 Response.End  
 end if  
 Response.AddHeader "Content-Disposition", "attachment; filename=" & myfile.name  
 Response.AddHeader "Content-Length", intFilelength  
 Response.CharSet = "UTF-8" 			'设置编码格式
 Response.ContentType = "application/octet-stream"  
 Response.BinaryWrite Objstream.Read    '向HTTP输出写入二进制信息
 Response.Flush 						'将缓冲信息发送给浏览器
 Objstream.Close  
 Set Objstream = Nothing  
 End Function  
 %> 
</body>
</html>