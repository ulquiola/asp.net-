<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
<style type="text/css">
<!--
.STYLE15 {
	color: #FFFFFF;
	font-size: 9pt;
}
body {
	margin-top: 0px;
}
.STYLE17 {
	font-size: 9pt;
	color: #000000;
}
.STYLE18 {
	font-size: 10pt;
	font-weight: bold;
}
-->
</style>
</head>
<body>
<table width=380 border=1 align=center cellpadding=0 cellspacing=0 bordercolor=#000000>
  <tr>
    <td width="376" height="19" valign=middle bordercolor=#FFFFFF bgcolor=#A90303>    <div align="center" class="STYLE15 STYLE18">订 单 信 息 查 询</div></td>
  </tr>
  <tr>
    <td valign=top bordercolor=#FFFFFF bgcolor=#FFFFFF>&nbsp;</td>
  </tr>
  <tr>
    <td height="89" valign=top bordercolor=#FFFFFF bgcolor=#FFFFFF>
      <div align=center>
	  <script language="javascript">
	  	function get()							//创建自定义函数
			{
				if(form1.OrderID.value=="")		//判断定单号是否为空
					{alert("请输入订单号码");	//弹出提示对话框
					form1.OrderID.focus();		//获取焦点
					}
				if(isNaN(form1.OrderID.value))	//判断订单号是否是数字型
				{
				alert("订单号码必须为数字的?");	//弹出提示对话框
				form1.OrderID.focus();			//获取焦点
				}
				else
					{
					window.close();				//关闭窗口
					window.open("order_index.asp?OrderID="+form1.OrderID.value);	//打开指定的页面
					}
			}			
	  </script>
        <form action="" method="post" name="form1">
          <span class="STYLE17">请输入订单号码：</span>
          <input name="OrderID" type="text" id="OrderID">
          <br>
            <br>
            <input name="Submit" type="submit" class="go-wenbenkuang" value="查询" onClick="get()">            
         	&nbsp;&nbsp;&nbsp;&nbsp;
	 	  <input name="Submit2" type="button" class="go-wenbenkuang" value="返回" onclick="javascript:window.close()"> 
            <br>
        </form>
      </div>
      <div align="right"></div>
    <div align="right"></div></td>
  </tr>
</table>
</body>
</html>
