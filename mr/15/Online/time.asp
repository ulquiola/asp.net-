<!--#include file=function.asp-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<link href="CSS/css.css" rel="stylesheet" type="text/css" />
<script src="js/jsc.js" type="text/javascript"></script>

</head>

<body>
<table width="1000" height="285" border="0">
  <tr>
    <td width="160" valign="top" class="rightLine">
	<!--#include file=left.asp-->	</td>
    <td width="758" valign="top"><div>首页   -&gt; 时段分析</div>
      <div>
        <table width="90%">
            <form name="changedate" id="changedate" onsubmit="return checkDate();">
              <tbody>
                <tr>
                  <td>
				  <div class="Title">时间段选择</div>
                    从
                    <input id="startdate"  size="10" value="<%=date()%>" name="txtStartdate" />
                    到
                    <input id="enddate"  size="10" value="<%=date()%>" name="txtEnddate" />
                    <input name="" type="button" onclick="GoUrl('select','time.asp');" value="选择时间段"/>
                    <input name="button" type="button" onclick="GoUrl('today','time.asp');" value="今日统计" />
                    <input name="button" type="button" onclick="GoUrl('yesterday','time.asp');" value="昨日统计" />
                    <input name="button" type="button" onclick="GoUrl('month','time.asp');" value="当月统计" />
					
                    <input name="button" type="button" onclick="GoUrl('30','time.asp');" value="最近30天" /></td>
                </tr>
              </tbody>
            </form>
        </table>
    </div>
	
	<%DayVisit()%>
	
	<%HoursVisit%>
	 </tr>
</table>
</body>
</html>
