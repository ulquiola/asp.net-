<!--#include file="Conn/conn.asp" -->
<%
if request.Form("aa")<>"" then
	files=request.Form("aa")
end if
set rs=conn.execute("select * from tb_talk")
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
	var table=document.all.pay;
	row=table.rows.length;
	column=table.rows(1).cells.length;
	var excelapp=new ActiveXObject("Excel.Application");
	excelapp.visible=true;
	objBook=excelapp.Workbooks.Add(); //����µĹ�����
	var objSheet = objBook.ActiveSheet;
	for(i=0;i<row;i++){
		for(j=0;j<column;j++){
		  objSheet.Cells(i+1,j+1).value=table.rows(i).cells(j).innerHTML.replace("&nbsp;","");
		}
	}
	objBook.SaveAs("<%=files%>payList.xls",1);
	excelapp.DisplayAlerts=false;
	excelapp.Workbooks.Close;
	excelapp.Quit();
	excelapp=null;
	setTimeout(CollectGarbage,1);		//�������ܹؼ������û��������䣬��ʹExcel���ڹر��ˣ����ǽ����л����ڣ����Ա������ļ����ܴ�
</script>