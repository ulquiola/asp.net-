<!--#include file="conn.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title></title>
<link href="CSS/index.css" rel="stylesheet"/>
<link href="CSS/scrip.css" rel="stylesheet"/>
<script language="javascript">
var iLayerMaxNum=<%=ss+100%>;
</script>
<script language="javascript" src="JS/AjaxRequest.js"></script>
<script language="javascript" src="JS/index.js"></script>
<script language="javascript" src="JS/add.js"></script>
<style type="text/css">
<!--
body {
	background-image: url(images/bg.jpg);
}
-->
</style>
</head>
<body>
<div style="display:none;" id="shadeDiv" onClick="Hide();"></div>
<!--#include file="Top.asp" -->
<div id="ntitle">����10����Ը��</div>
<div id="nmarquee">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
<td>
<table width="100%" height="25" border="0" cellpadding="0" cellspacing="0" background="images/bg_search1.gif">
<tr>
<td>
<marquee onMouseOver="this.stop()" onMouseOut="this.start()" scrollamount="2">
<%
Set rs_1=Server.CreateObject("ADODB.RecordSet")	
sql_1="Select top 10 * From tb_wish Order By id DESC"
rs_1.Open sql_1,Conn,1,3
while not rs_1.eof
%>
<strong><font color="#333366"><%=rs_1("wellWisher")%></font></strong>
<font color="#FF3366">&nbsp;ף��&nbsp;</font>
<strong><font color="#333366"><%=rs_1("wishMan")%></font></strong>��<%=rs_1("Content")%>
<%
rs_1.movenext
wend
%> 
</marquee>
</td>
       </tr>
      </table>
	 </td>
  </tr>
</table>
<script language="javascript">
function createNewScrip(id){	//ʵʱ��ʾ�ո���ӵ�����
	var newScrip="<div id='scrip"+id+"' class='Style"+scripForm.color.value+"' style='left:100px;top:200px;z-index:200' onmousedown='Move(this,event)' ondblclick=\"Show("+id+",'shadeDiv')\"><p class='Num'>������ţ�"+id+"&nbsp;&nbsp;������<span id='hitsValue"+id+"'>0</span><img src='images/close.gif' alt='�ر�' onclick='myClose("+id+")'></p><br /><p class='Detail'><img src='images/face/face_"+scripForm.face.value+".gif'><span class='wishMan'>"+scripForm.wishMan.value+"</span>	<br />"+scripForm.content.value+"</p><p class='wellWisher'>"+scripForm.wellWisher.value+"</p><p class='comment'><a href='#' onclick='holdout("+id+")'>[֧��]</a></p><p class='Date'>"+getTime()+"</p></div>";
	document.getElementById("main").innerHTML=document.getElementById("main").innerHTML+newScrip;
}
</script>
</div>
</div>
<div id="main">
<%
set rs=server.CreateObject("adodb.recordset")
sql="select * from tb_wish"
rs.open sql,conn,1,3
ss=rs.recordcount
while not rs.eof
i=i+1
randomize
T = Int(310*Rnd+255)
L = Int(720*Rnd+10)
%>
<div id="scrip<%=rs("id")%>" class="Style<%=rs("color")%>" style="left:<%=l%>px;top:<%=t%>px; z-index:<%=i%>" onMouseDown="Move(this,event)" onDblClick="Show(<%=rs("id")%>,'shadeDiv')">
		<p class='Num'>������ţ�<%=rs("id")%>&nbsp;&nbsp;������<span id="hitsValue<%=rs("id")%>"><%=rs("hits")%></span><img src='images/close.gif' alt='�ر�' onClick="myClose(<%=rs("id")%>)"></p>
		<br /> 
		<p class="Detail">
		<img src="images/face/face_<%=rs("face")%>.gif">
		<span class='wishMan'><%=rs("wishMan")%></span>
		<br /><%=rs("content")%></p>
		<p class='wellWisher'><%=rs("wellWisher")%></p>
		<p class="comment"><a href="#" onClick="holdout(<%=rs("id")%>)">[֧��]</a></p>
		<p class="Date"><%=rs("sendTime")%></p>
</div>
<%
rs.MoveNext
wend
rs.Close
Set rs=Nothing
Conn.Close
Set Conn=Nothing
%>
</div>
<div id="footer"></div>
<%=Bottom%>
<div id="notClickDiv"></div>
<script language="javascript">
 function myReload(){
	document.getElementById("sss").src=document.getElementById("sss").src+"?nocache="+new Date().getTime();
 }
</script>
<div id="scrip_add">
  <table width="100%" height="300" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td width="15" height="14" background="images/scrip_add_leftTop.gif"></td>
      <td height="14" background="images/scrip_add_Top.gif"></td>
      <td width="15" background="images/scrip_add_rightTop.gif"></td>
    </tr>
    <tr>
      <td height="272" background="images/scrip_add_Left.gif"></td>
      <td valign="top" bgcolor="#FFFFFF">
	  <form action="" method="post" name="scripForm">
	  <table width="100%" height="188" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="16%" align="center">ף������</td>
          <td width="84%" align="left" colspan="3">
             <input type="text" name="wishMan" size="30"  onkeyup="javascript:InputInfo(this,'pWishMan');"/>*</td>
        </tr>
        <tr>
          <td align="center">ף �� �ߣ�</td>
          <td align="left" colspan="3">
            <input type="text" name="wellWisher" onKeyUp="javascript:InputInfo(this,'pWellWisher');" size="30" value="����" onClick="if(this.value =='����') this.value=''; "/> * </td>
   		</tr>
		<tr>
		   <td align="center">������ɫ��</td>
		   <td height="50px" align="left" colspan="3"><span class="color0" onClick="ColorChoose('0')"></span> <span class="color1" onClick="ColorChoose('1')"></span> <span class="color2" onClick="ColorChoose('2')"></span> <span class="color3" onClick="ColorChoose('3')"></span> <span class="color4" onClick="ColorChoose('4')"></span><span class="color5" onClick="ColorChoose('5')"></span><input type="hidden" id="color" name="color" value="0"/></td>
		</tr>
		<tr>
			  <td height="65px" rowspan="2" align="center">����ͼ����</td>
			  <td height="24" colspan="3" align="left">��ѡ������ͼ����� 
			  <a href="#" onClick="face_1.style.display='block';face_2.style.display='none';face_3.style.display='none';">��֮��</a>
			  <a href="#" onClick="face_1.style.display='none';face_2.style.display='block';face_3.style.display='none';">��֮��</a>
			  <a href="#" onClick="face_1.style.display='none';face_2.style.display='none';face_3.style.display='block';">�λ�ˮ��</a>
			  </td>
		 </tr>
		 <tr>
			  <td align="left" colspan="3">
			  <div id="face_1" style="display:block">
			  <%for i=0 to 5%>
				 <img src="images/face/face_<%=i%>.gif" width="56" height="56" onClick="javascript:faceChoose('<%=i%>');" /><%next%>  
			  </div>
			  <div id="face_2" style="display:none">
			  <%for i=6 to 11%>
				 <img src="images/face/face_<%=i%>.gif" width="56" height="56" onClick="javascript:faceChoose('<%=i%>');" /><%next%>
			  </div>
			   <div id="face_3" style="display:none">
			  <%for i=12 to 17%>
				 <img src="images/face/face_<%=i%>.gif" width="56" height="56" onClick="javascript:faceChoose('<%=i%>');" /><%next%>  
			  </div>
			  <input type="hidden" value="0" name="face" id="face"/></td>
		 </tr>
		 <tr>
			  <td align="center">�������ݣ�</td>
			  <td align="left" colspan="3"> <textarea name="content" id="content" rows="7" cols="49" class="wenbenkuang" onKeyDown="CountStrByte(this,this.form.total,this.form.used,this.form.remain);"onkeyup="CountStrByte(this,this.form.total,this.form.used,this.form.remain);" ></textarea> * </td>
		 </tr>
		 <tr>
			<td height="33" align="center" style="padding-left:10px">��&nbsp;&nbsp;�ڣ�</td>
			<td style="padding-left:10px" colspan="3">������� 
			<input name="total" type="text" disabled class="noborder" id="total" value="200" size="4"> 
			���ֽ� �����ֽڣ�&nbsp;
			<input name="used" type="text" disabled class="noborder" id="used"  value="0" size="4">                        
			ʣ���ֽڣ�
			<input name="remain" type="text" disabled class="noborder" id="remain" value="200" size="4">
				</td>
			  </tr>
			<tr>
			  <td align="center" valign="bottom">�� ֤ �룺</td>
			  <td width="300px" align="left">
			<input id="yanzheng" name="yanzheng" type="text" value="" size="6" maxlength="4" onFocus="this.value='';" onpaste="var tt=clipboardData.getData('text'); if(!/\D/.test(tt))value=tt.replace(/^0*/,''); return false;" onKeyUp="if(/(^0+)/.test(value))value=value.replace(/^0*/, '')" onKeyPress="return event.keyCode>=48 && event.keyCode<=57;"><span class="red">*</span> �������ұ����֣�<input id="chknm" type="hidden" /><img src="" id="sss"><a href="#" onClick="showval()">&nbsp;������?��һ��</a>
			 </td>
				<td width="40px" valign="bottom">&nbsp;&nbsp;&nbsp;
				</td>
			</tr>
			<tr>
			  <td colspan="3" align="center">
			  <input type="button" id="btn_Submit" name="btn_Submit" class="btn_grey" value="����" onClick="scripSubmit()">&nbsp;
			  <input type="button" class="btn_grey" value="�ر�" onClick="close_window()"/>
			  </td>
			</tr>
      </table>
	  </form>
	  </td>
      <td background="images/scrip_add_Right.gif">&nbsp;</td>
    </tr>
    <tr>
      <td height="14" background="images/scrip_add_leftBottom.gif"></td>
      <td height="14" background="images/scrip_add_Bottom.gif"></td>
      <td background="images/scrip_add_rightBottom.gif"></td>
    </tr>
  </table>
</div>
<input type="hidden" value="<%=ss%>" id="hRsCount">
<!--����Ԥ��-->
<div id="preview"  class="StyleTemp">
		<p class='Num'>����Ԥ����</p>
		<br/> 
		<p class="Detail">
		<img src="images/face/face_0.gif" id="pFace">
		<span class='wishMan'><span class="wishMan" id="pWishMan"></span></span>
		<br />
		<span id="pContext"></span></p>
		<p class='wellWisher'><span class="wellWisher" id="pWellWisher">����</span></p>
</div>
</body>
</html>