<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!-- #include file="Connections/conn_login.asp" -->  <!--�������ݿ������ļ�-->
<% session.Timeout=120 
if request.Form("UserName")<>"" and request.Form("PWD")<>"" then
session("UserName")=request.Form("UserName")
session("PWD")=request.Form("PWD")
sql="select Name,PWD from DB_manager where name='" & session("UserName")&"'"
set rs=conn.execute(sql)
if rs.eof then %>
  <script language="javascript">
  alert("������Ĺ���Ա���ƴ������������룡");
  </script> 
<% else 
	if rs("PWD")=session("PWD") then  %>
     <script language="javascript">
      window.location.href="manager/default.asp"
  	 </script>
  <%else%>
   <script language="javascript">
  alert("������Ĺ���Ա����������������룡");
  history.back();
  </script>  		
	<%end if
end if
end if
%>
<script language="javascript">
function mycheck(){
if (form1.UserName.value=="")
{alert("���������Ա������");form1.UserName.focus();return;}
if(form1.PWD.value=="")
{alert("���������룡");form1.PWD.focus();return;}
form1.submit();
}
</script>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>���Ͽͻ�����ϵͳ--����Ա��¼��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="style.css" rel="stylesheet">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	background-image: url(Images/bg.gif);
}
.style1 {color: #FFFFFF}
-->
</style></head>
<body>
<table width="595" height="242" border="0" cellpadding="-2" cellspacing="-2" background="Images/login.gif">
  <tr>
    <td valign="top"><table width="595" height="399" cellpadding="-2" cellspacing="-2">
      <tr>
        <td width="112" height="140">&nbsp;</td>
        <td width="370">&nbsp;</td>
        <td width="111"></td>
      </tr>
      <tr>
        <td height="211">&nbsp;</td>
        <td valign="top"><div align="center">
          <form name="form1" method="POST" action="login_manager.asp">
            <table width="76%" height="154" border="1" cellpadding="-2"  cellspacing="-2"
			 bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#669999">
            <tr>
              <td height="19" bgcolor="#6B8E94"><div align="center" class="style1">
			  ==========&nbsp;[����Ա��¼]&nbsp;==========</div></td>
            </tr>
            <tr>
              <td height="93"><div align="center">
                <table width="63%" height="89" cellpadding="-2"  cellspacing="-2">
                  <tr>
				  <script language="javascript">
				  //��Ӧ�س��¼�
				  function EnterE(packName){
				    if(event.keyCode=="13")   //�ж��û����µ��Ƿ�Ϊ�س���
				      {packName.focus(); }��//��һ���ؼ���ý���
				   }
				  </script>
                    <td width="48%"><div align="center"><span class="style1">����Ա���ƣ�
					</span></div></td>
                    <td width="52%"><div align="left">
                      <input name="UserName" type="text" class="Sytle_text" id="UserName" 
					  size="20" onKeyPress="EnterE(document.form1.PWD)">
                    </div></td>
                  </tr>
                  <tr>
                    <td><div align="center"><span class="style1">����Ա���룺</span></div></td>
                    <td><div align="left">
                      <input name="PWD" type="password" class="Sytle_text" id="PWD" size="20"
					   onKeyPress="EnterE(document.form1.Button)">
					</div></td>
                  </tr>
                  <tr>
                    <td height="33" colspan="2"><div align="center">
                        <input name="Button" type="button" class="Style_button" value="��¼" 
						onClick="mycheck()">
                        <input name="Submit2" type="reset" class="Style_button" value="����">
                    </div></td>
                    </tr>
                </table>
              </div></td>
            </tr>
          </table>
		  </form>
        </div></td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
</html>
