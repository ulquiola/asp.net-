<%
SqlDatabaseName = "shop"            			      '���ݿ���
SqlUsername = "sa"                                    '�û���
SqlPassword = ""                                      '�û�����
Dim Conn,ConnStr
Set Conn = Server.CreateObject("ADODB.Connection")
ConnStr = "Provider = Sqloledb; Persist Security Info=false; User ID = " & SqlUsername & "; Password = " & SqlPassword & "; Initial Catalog = " & SqlDatabaseName & ";"
Conn.Open ConnStr
%>
<title>�����̳�--����ʡ���տƼ����޹�˾</title>
<script>top.document.title="�����̳�--����ʡ���տƼ����޹�˾";</script>
<style>
BODY {
	font-family: "����";
	font-size: 9pt;
	font-style: normal;
	line-height: 160%;
	color: #000000;
	background-color: #FFFFFF;
}
TABLE {
	font-family: "����";
	font-size: 9pt;
	font-style: normal;
}
A:link {
	FONT-SIZE: 12px; COLOR: #000000; TEXT-DECORATION: none
}
A:visited {
	FONT-SIZE: 12px; COLOR: #000000; TEXT-DECORATION: none
}
A:active {
	FONT-SIZE: 12px; COLOR: #215DC6; TEXT-DECORATION: none
}
A:hover {
	FONT-SIZE: 12px; COLOR: #215DC6; TEXT-DECORATION: none;position: relative; right: 0px; top: 1px
}
.style1 {color: #f2ab5b}
</style>