<!--#include file=function.asp-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
<link href="CSS/css.css" rel="stylesheet" type="text/css" />
<script src="js/jsc.js" type="text/javascript"></script>
</head>

<body>
<table width="1000" height="285" border="0">
  <tr>
    <td width="160" valign="top" class="rightLine">
	<!--#include file=left.asp-->	</td>
    <td width="758" valign="top">
	<div>��ҳ   -&gt; �ܷ�ҳ����ʷ</div>
      <div>
        <table width="90%">
            <form name="changedate" id="changedate" onsubmit="return checkDate();">
              <tbody>
                <tr>
                  <td>
				  <div class="Title">ʱ���ѡ��</div>
                    ��
                    <input id="startdate"  size="10" value="<%=date()%>" name="txtStartdate" />
                    ��
                    <input id="enddate"  size="10" value="<%=date()%>" name="txtEnddate" />
             <input name="" type="button" onclick="GoUrl('select','pageComeHistory.asp?pageurl=<%=server.URLEncode(pageUrl)%>');" value="ѡ��ʱ���"/>
                    <input name="button" type="button" onclick="GoUrl('today','pageComeHistory.asp?pageurl=<%=server.URLEncode(pageUrl)%>');" value="����ͳ��" />
                    <input name="button" type="button" onclick="GoUrl('yesterday','pageComeHistory.asp?pageurl=<%=server.URLEncode(pageUrl)%>');" value="����ͳ��" />
                    <input name="button" type="button" onclick="GoUrl('month','pageComeHistory.asp?pageurl=<%=server.URLEncode(pageUrl)%>');" value="����ͳ��" />
					
                    <input name="button" type="button" onclick="GoUrl('30','pageComeHistory.asp?pageurl=<%=server.URLEncode(pageUrl)%>');" value="���30��" /></td>
                </tr>
              </tbody>
            </form>
        </table>
    </div>
	
	<%PageComeHistory()%>
	
	</td>
  </tr>
</table>
</body>
</html>
