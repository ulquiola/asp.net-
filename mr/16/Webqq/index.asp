<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
</head>

<body>
<script language="javascript">
	window.open("index1.asp","newwindow","height=199,width=300,toolbar=no,menubar=no,scrollbars=no,resizable=no,location=no,status=no")//弹出一个新窗口，并对该窗口的高、宽进行设定。同时去掉浏览器工具条、菜单条、滚动条、地址栏、状态栏、设置固定窗口
	window.opener=null;//表示对父窗口不进行任何操作
	window.close();//关闭窗口
</script>
</body >
</html>
