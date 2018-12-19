<!--#include file="Conn/conn.asp" -->
<%
if request.Form("aa")<>"" then					'判断文件路径是否为空
	files=request.Form("aa")					'为files变量赋值
end if									
set rs=conn.execute("select * from tb_talk")	'执行查询语句
%>
<table   border="1" align="center" cellspacing="0" id="pay">
  <tr>
  <%
  j=2
  for i=0 to rs.fields.count-1
  %>
    <td width="80" height="32" align="center" bgcolor="#FFFFCC"><%=rs.fields(i).name%></td>
	  <%next%>
  </tr>
  <%do while not rs.eof%>
  <tr>
  <%for i=0 to rs.fields.count-1%>
    <td height="29" align="center" ><%=rs(i)%></td>
	<%next%>
  </tr>
  <%rs.movenext
  j=j+1
  loop
  rs.close%>
</table>
<script language="javascript">
	var table=document.all.pay;			//声明变量
	row=table.rows.length;				//获取表格行的宽度
	column=table.rows(1).cells.length;	//获取单元格的宽度
	var excelapp=new ActiveXObject("Excel.Application");	//启动Excel
	excelapp.visible=true;									//设置工作簿可见
	objBook=excelapp.Workbooks.Add(); 						//添加新的工作簿
	var objSheet = objBook.ActiveSheet;						//激活工作簿
	for(i=0;i<row;i++){										
		for(j=0;j<column;j++){
		  objSheet.Cells(i+1,j+1).value=table.rows(i).cells(j).innerHTML.replace("&nbsp;","");
		  //获取行和单元格数
		}
	}
	objBook.SaveAs("<%=files%>payList.xls",1);//将生成的Excel自动保存到指定路径下
	excelapp.DisplayAlerts=false;				//隐藏活跃的表格
	excelapp.Workbooks.Close;					//关闭工作簿
	excelapp.Quit();							//退出Excel
	excelapp=null;								//清空变量值
	setTimeout(CollectGarbage,1);		//这条语句很关键，如果没有这条语句，即使Excel窗口关闭了，但是进程中还存在，所以保存后的文件不能打开
	alert("聊天记录信息导出成功！");			//弹出提示对话框
	window.close();								//关闭窗口
</script>