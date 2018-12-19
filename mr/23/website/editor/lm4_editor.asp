<%@LANGUAGE="VBSCRIPT"%>
<% Session.Timeout=60 %>
<!--#include file="../Conn/conn.asp" -->
<%
Dim r__MRColParam
r__MRColParam = "1"
if (Session("MR_username") <> "") then r__MRColParam = Session("MR_username")
%>
<%
set r = Server.CreateObject("ADODB.Recordset")
r.ActiveConnection = MR_website_STRING
r.Source = "SELECT * FROM website WHERE id = '" + Replace(r__MRColParam, "'", "''") + "'"
r.CursorType = 0
r.CursorLocation = 2
r.LockType = 3
r.Open()
r_numRows = 0
%>
<HTML>
<HEAD>
<TITLE>建站平台</TITLE>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="styleMain.css">
<script language="JavaScript" src="HR_Select.js"></script>
<script language=JavaScript>
		function openWindow(url,name,width,height) {
        		var Win = window.open(url,name,'width=' + width + ',height=' + height + ',resizable=1,scrollbars=no,menubar=yes,status=yes' );
		}

		function AdminTestCookies() {
			if( !document.cookie ) {
				alert("系统登陆需要 Cookies 支持，请开启您的 Cookies 功能，然后再试！");
				return false;
			}else{
				document.Loginfrm.submit();
			}
		}
	</script>
</HEAD>
<body bgcolor=#CCCCCC topmargin=1 marginheight=5 leftmargin=4 marginwidth=5 oncontextmenu=self.event.returnValue=false>
<form name="HtmlEditForm" action="lm4_save.asp" method=post>
  <script src="htmleditor/forEdit.js" language="JScript"></script>
  <table border="1" align="center" bordercolor="#CCCCCC" cellpadding="0" cellspacing="0" width="100%" height="400">
    <tr>
      <td colspan="2" align="center"><input type=hidden name=body id=Hidden1>
        <input type="hidden" name="content">
        <textarea name="oldContent" style="display:none"><%=(r.Fields.Item("main4").Value)%></textarea>
        <script language="JavaScript">
<!--
	strHTML=document.all["oldContent"].value;

	// 主要是为了颜色配置而设定的
	if (document.createAttribute) {
		document.write("<object id=\"dlgColHelper\" classid=\"clsid:3050f819-98b5-11cf-bb82-00aa00bdce0b\" width=\"0px\" height=\"0px\"></object>");
	}

	if((window.navigator.userAgent.indexOf("IE 5") > 1) || (window.navigator.userAgent.indexOf("IE 6") > 1)) {
              document.write('<iframe ID="editor" NAME="editor" style="width: 100%; height:400;display:none"></iframe>');
	}


	function switchstatus(flag){
		document.frames.editor.frames.edit1.swapModes(flag);
		var i;
		for(i=1;i<4;i++){
			document.all["status" + i].style.display = "none";
		}
		document.all["status" + flag].style.display = "block";
	}

	function winhidden(){
		if (confirm("您确信要放弃所有改动退出吗？\n退出后将无法保存您所做的改动。")){
			document.all.editor.src = "";
			window.close();
		}
	}

	function win_init(){
		document.all.editor.src = "htmleditor/edit.htm";		
		window.status = "程序载入中，请等待……";
	}

	function save(){
		document.frames.editor.savefile();
	}

	window.onload = win_init
      
	function UploadComplete(URL){
		if ((URL != null) && (URL != "")) {
			if (URL.indexOf(":") == -1)
				doFormat("InsertImage", "http://" + URL);
			else			
				doFormat("InsertImage", URL);
		}
		document.all.UploadImg.style.display = "none";
		document.forms["upload"].reset();
	}
//-->
</script>
      </td>
    </tr>
  </table>
  <table width="100%" cellpadding="0" cellspacing="0" border="0">
    <tr>
      <td colspan="2" height="1" bgcolor="#FFFFFF"></td>
    </tr>
    <tr>
      <td colspan="2" height="2" bgcolor="#999999"></td>
    </tr>
    <tbody>
      <tr>
        <td colspan="2" height="60" align="center"><input type="hidden" name="userId" value="xingworld">
          <input type="submit" name="submit" style="background-color=buttonface" value="  提  交  " onClick="save(); document.HtmlEditForm.content.value = document.HtmlEditForm.body.value;">
          &nbsp;&nbsp;
          <input type="button" name="submit2" style="background-color=buttonface" value="  放  弃  " onClick="window.close();return false;">
        </td>
      </tr>
    </tbody>
  </table>
</form>
</body>
</html>
<%
r.Close()
%>
