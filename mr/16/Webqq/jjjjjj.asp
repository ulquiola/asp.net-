<!--#include file="Conn/conn.asp" -->
<%
if request.Form("aa")<>"" then					'�ж��ļ�·���Ƿ�Ϊ��
	files=request.Form("aa")					'Ϊfiles������ֵ
end if									
set rs=conn.execute("select * from tb_talk")	'ִ�в�ѯ���
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
	var table=document.all.pay;			//��������
	row=table.rows.length;				//��ȡ����еĿ��
	column=table.rows(1).cells.length;	//��ȡ��Ԫ��Ŀ��
	var excelapp=new ActiveXObject("Excel.Application");	//����Excel
	excelapp.visible=true;									//���ù������ɼ�
	objBook=excelapp.Workbooks.Add(); 						//����µĹ�����
	var objSheet = objBook.ActiveSheet;						//�������
	for(i=0;i<row;i++){										
		for(j=0;j<column;j++){
		  objSheet.Cells(i+1,j+1).value=table.rows(i).cells(j).innerHTML.replace("&nbsp;","");
		  //��ȡ�к͵�Ԫ����
		}
	}
	objBook.SaveAs("<%=files%>payList.xls",1);//�����ɵ�Excel�Զ����浽ָ��·����
	excelapp.DisplayAlerts=false;				//���ػ�Ծ�ı��
	excelapp.Workbooks.Close;					//�رչ�����
	excelapp.Quit();							//�˳�Excel
	excelapp=null;								//��ձ���ֵ
	setTimeout(CollectGarbage,1);		//�������ܹؼ������û��������䣬��ʹExcel���ڹر��ˣ����ǽ����л����ڣ����Ա������ļ����ܴ�
	alert("�����¼��Ϣ�����ɹ���");			//������ʾ�Ի���
	window.close();								//�رմ���
</script>