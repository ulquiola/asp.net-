<!--#include file="../include/md5.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<%
sub updateuser()	'�޸��û���Ϣ
	sql="select * from [user] where id="&request("id")&" and name='"&trim(session("user"))&"';"
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,3,3
	rs("mail")=trim(request("mail"))
	rs("youbian")=trim(request("youbian"))
	rs("xingming")=trim(request("xingming"))
	rs("shenfenzheng")=trim(request("shenfenzheng"))
	rs("tel")=trim(request("tel"))
	rs("qq")=trim(request("qq"))
	rs("tishi")=trim(request("tishi"))
	rs("huida")=trim(request("huida"))
	rs("dizhi")=trim(request("dizhi"))
	rs.update
	rs.close
	set rs=nothing
	response.Write("<script>alert('�޸ĳɹ���');window.location.href='usercenter.asp?action=1';</script>")
end sub

sub updatepass()	'�޸��û�����
	sql="select * from [user] where id="&request("id")&" and name='"&trim(session("user"))&"';"	'������ѯ
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,3,3
	rs("pass")=md5(request("newpass1"))	'���û��ύ������ͨ��MD5���м���
	rs.update
	rs.close
	set rs=nothing
	response.Write("<script>alert('�޸ĳɹ���');window.location.href='usercenter.asp?action=1';</script>")
end sub 

sub thispass()	'�û������һ�
	sql="select * from [user] where name='"&request("name")&"';"	'��ȷ���û�������û����Ƿ����
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,3,3
	if not rs.eof then	'�������
		if rs("huida")=request("huida") and rs("tishi")=request("tishi") then	'�ж��û������������ʾ����������ش������ݿ����Ƿ�һ��
			session("shijian")=rs("shijian2")	'�����û��Ѿ���½
			session("cishu")=rs("cishu")
			session("user")=trim(request("name"))
			rs("shijian2")=now()
			rs("cishu")=rs("cishu")+1
			rs.update
			response.Write("<script>alert('ϵͳ��ʾ\n\n��֤ͨ���������Զ���½״̬��');window.location.href='usercenter.asp?action=3';</script>")	'�����û����ĵ������޸�ģ����
			response.End()
		end if
	end if
	rs.close
	set rs=nothing
	response.Write("<script>alert('���ݲ���ȷ���޷���ѯ��');window.location.href='usercenter.asp?action=1';</script>")	'Ҳ�����û�������ʾ���ش����κ�һ�����벻����û���������ж������ִ�����Ϊ�����Ը����û�������ϸ����Ϣ���������޵Ĳ½⣩
end sub
%>

<%sub	lookuser()	'�鿴�û���Ϣ%>
<table width="96%"  border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="BDBDBC">
<%
sql="select * from [user] where name='"&trim(session("user"))&"';"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
if not rs.eof then
%>
<form name="myform" action="lookuser.asp" method="post">  <tr> 
    <td align="center"><span class="style5">��ϸ��Ϣ</span></td>
  </tr>
  <tr>
    <td valign="top" bgcolor="#FFFFFF"><br> 
      <table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="BDBDBC">
        <tr height="20" bgcolor="#FFFFFF" align="center"> 
          <td width="21%" height="30">��&nbsp;��&nbsp;����</td>
          <td width="79%"><div align="left"><%=rs("name")%></div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td height="30">�����ʼ���</td>
          <td><div align="left"><%=rs("mail")%></div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td height="30">��&nbsp;&nbsp;&nbsp;&nbsp;����</td>
          <td><div align="left"><%=rs("xingming")%></div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td height="30">��&nbsp;&nbsp;&nbsp;&nbsp;����</td>
          <td><div align="left"><%=rs("tel")%></div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td height="30">��&nbsp;��&nbsp;֤��</td>
          <td><div align="left"><%=rs("shenfenzheng")%></div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td height="30">��&nbsp;&nbsp;&nbsp;&nbsp;�ࣺ</td>
          <td><div align="left"><%=rs("youbian")%></div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td height="30">��&nbsp;&nbsp;&nbsp;&nbsp;ַ��</td>
          <td><div align="left"><%=rs("dizhi")%></div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td height="30">��ϵ&nbsp;&nbsp;QQ��</td>
          <td><div align="left"><%=rs("qq")%></div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td height="30">��&nbsp;&nbsp;&nbsp;&nbsp;ʾ��</td>
          <td><div align="left"><%=rs("tishi")%></div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td height="30">ע��ʱ�䣺</td>
          <td><div align="left"><%=rs("shijian1")%>
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td height="30">�ϴε�¼��</td>
          <td><div align="left"><%=rs("shijian2")%>
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td height="30">��¼������</td>
          <td><div align="left"><%=rs("cishu")%>
          </div></td>
        </tr>
      </table>
      <br></td>
  </tr>
</form>
<%
end if
rs.close
set rs=nothing
%>
</table>
		  </td>
        </tr>
</table>
<%end sub%>

<%sub upuser()	'�޸��û���Ϣ��%>
<script>
function chk()
{
	if (document.myform.mail.value=="")
	{
		document.myform.mail.focus();
		alert("����������ʼ���");
		return false;
	}
	
	if (document.myform.youbian.value=="")
	{
		document.myform.youbian.focus();
		alert("�������ʱ࣡");
		return false;
	}
	
	if (document.myform.xingming.value=="")
	{
		document.myform.xingming.focus();
		alert("��������ʵ������");
		return false;
	}
	
	if (document.myform.tel.value=="")
	{
		document.myform.tel.focus();
		alert("��������ϵ�绰��");
		return false;
	}
	
	if (document.myform.shenfenzheng.value=="")
	{
		document.myform.shenfenzheng.focus();
		alert("���������֤��");
		return false;
	}
	
	if (document.myform.dizhi.value=="")
	{
		document.myform.dizhi.focus();
		alert("�������ַ��");
		return false;
	}
	
	if (document.myform.qq.value=="")
	{
		document.myform.qq.focus();
		alert("��������ϵqq��");
		return false;
	}
	
	if (document.myform.tishi.value=="")
	{
		document.myform.tishi.focus();
		alert("������������ʾ��");
		return false;
	}
	
	if (document.myform.huida.value=="")
	{
		document.myform.huida.focus();
		alert("����������ش�");
		return false;
	}
	
}
</script>
<table width="96%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td>
      <table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="BDBDBC">
<%
sql="select * from [user] where name='"&trim(session("user"))&"';"
'response.Write ("����Ӧ����ʾ�û�����"&trim(session("user"))&"")
'response.Write("<br>")
'response.Write ("������ SQL ��䣺"&sql&"")
'response.End()	'ֹͣ����Ĳ���
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
if not rs.eof then
%>
        <form name="myform" action="usercenter.asp" method="post"  onSubmit="return chk();">
          <tr>
            <td align="center"><span class="style5">�޸���Ϣ</span></td>
          </tr>
          <tr>
            <td valign="top" bgcolor="#FFFFFF"><br>
                <table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="BDBDBC">
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td width="21%">��&nbsp;��&nbsp;����</td>
                    <td width="79%"><div align="left"><%=rs("name")%></div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td>�����ʼ���</td>
                    <td><div align="left">
                        <input name="mail" type="text" value="<%=rs("mail")%>" size="30">
                    </div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td>��&nbsp;&nbsp;&nbsp;&nbsp;�ࣺ</td>
                    <td><div align="left">
                        <input name="youbian" type="text" value="<%=rs("youbian")%>" size="30">
                    </div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td>��&nbsp;&nbsp;&nbsp;&nbsp;����</td>
                    <td><div align="left">
                        <input name="xingming" type="text" value="<%=rs("xingming")%>" size="30">
                    </div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td>��&nbsp;&nbsp;&nbsp;&nbsp;����</td>
                    <td><div align="left">
                        <input name="tel" type="text" value="<%=rs("tel")%>" size="30">
                    </div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td>��&nbsp;��&nbsp;֤��</td>
                    <td><div align="left">
                        <input name="shenfenzheng" type="text" value="<%=rs("shenfenzheng")%>" size="30">
                    </div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td>��&nbsp;&nbsp;&nbsp;&nbsp;ַ��</td>
                    <td><div align="left">
                        <input name="dizhi" type="text" value="<%=rs("dizhi")%>" size="30">
                    </div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td>��ϵ&nbsp;&nbsp;QQ��</td>
                    <td><div align="left">
                        <input name="qq" type="text" value="<%=rs("qq")%>" size="30">
                    </div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td>��&nbsp;&nbsp;&nbsp;&nbsp;ʾ��</td>
                    <td><div align="left">
                        <input name="tishi" type="text" value="<%=rs("tishi")%>" size="30">
                    </div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td>��&nbsp;&nbsp;&nbsp;&nbsp;��</td>
                    <td><div align="left">
                        <input name="huida" type="password" value="<%=rs("huida")%>" size="30">
                    </div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td>ע��ʱ�䣺</td>
                    <td><div align="left"><%=rs("shijian1")%> </div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td>�ϴε�½��</td>
                    <td><div align="left"><%=rs("shijian2")%> </div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td>��½������</td>
                    <td><div align="left"><%=rs("cishu")%> </div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td colspan="2">
                      <input name="action" type="hidden" value="updateuser">
                      <input name="id" type="hidden" value="<%=rs("id")%>">
                      <input type="reset" name="reset" value="�ر�" onClick='javascript:window.close();'>
&nbsp;&nbsp;
                    <input name="submit" type="submit" value="�޸�"></td>
                  </tr>
                </table>
                <br></td>
          </tr>
        </form>
<%
end if
rs.close
set rs=nothing
%>
      </table>
    </td>
  </tr>
</table>
<%end sub%>

<%sub uppass()	'�޸��û������%>
<script>
function chk()
{
	if (document.myform.newpass1.value=="")
	{
		document.myform.newpass1.focus();
		alert("���������룡");
		return false;
	}
	
	if (document.myform.newpass2.value=="")
	{
		document.myform.newpass2.focus();
		alert("��ȷ�������룡");
		return false;
	}
	
	if (document.myform.newpass1.value!=document.myform.newpass2.value)
	{
		document.myform.newpass2.focus();
		alert("������������벻һ����");
		return false;
	}
}
</script>
<table width="96%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td>
      <table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="BDBDBC">
<%
sql="select * from [user] where name='"&trim(session("user"))&"';"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
if not rs.eof then
%>
        <form name="myform" action="usercenter.asp" method="post" onSubmit="return chk();">
          <tr>
            <td align="center"><span class="style5">�޸�����</span></td>
          </tr>
          <tr>
            <td valign="top" bgcolor="#FFFFFF"><br>
                <table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="BDBDBC">
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td width="21%">��&nbsp;��&nbsp;����</td>
                    <td width="79%"><div align="left"><%=trim(session("user"))%></div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td>ԭ&nbsp;��&nbsp;�룺</td>
                    <td><div align="left">
                        <input name="pass" type="password" value="<%=rs("pass")%>" size="30">
                    </div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td>�� �� �룺</td>
                    <td><div align="left">
                        <input name="newpass1" type="password" value="" size="30">
                    </div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td>ȷ�����룺</td>
                    <td><div align="left">
                        <input name="newpass2" type="password" value="" size="30">
                    </div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td colspan="2">
                      <input name="action" type="hidden" value="updatepass">
                      <input name="id" type="hidden" value="<%=rs("id")%>">
                      <input name="up" type="hidden" value="<%=request("up")%>">
&nbsp;&nbsp;
                    <input name="submit" type="submit" value="�޸�"></td>
                  </tr>
                </table>
                <br></td>
          </tr>
        </form>
<%
end if
rs.close
set rs=nothing
%>
      </table>
    </td>
  </tr>
</table>
<%end sub%>

<%sub lookpass()	'�����һ��ύ��%>
<script>
function chk()
{
	if (document.myform.name.value=="")
	{
		document.myform.name.focus();
		alert("�������û�����");
		return false;
	}
	
	if (document.myform.tishi.value=="")
	{
		document.myform.tishi.focus();
		alert("������������ʾ��");
		return false;
	}
	
	if (document.myform.huida.value=="")
	{
		document.myform.huida.focus();
		alert("����������ش�");
		return false;
	}
}
</script>
<table width="96%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td>
      <table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="BDBDBC">
        <form name="myform" action="usercenter.asp" method="post" onSubmit="return chk();"<%'��֤�Ƿ��Ѿ���д��������%>>
          <tr>
            <td align="center"><span class="style5">ȡ������</span></td>
          </tr>
          <tr>
            <td valign="top" bgcolor="#FFFFFF"><br>
                <table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="BDBDBC">
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td width="21%">��&nbsp;��&nbsp;����</td>
                    <td width="79%"><div align="left"><input type="text" name="name" value=""></div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td width="21%">������ʾ��</td>
                    <td width="79%"><div align="left"><input name="tishi" type="text" value="" size="30"></div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td>����ش�</td>
                    <td><div align="left">
                        <input name="huida" type="password" value="" size="30">
                    </div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td colspan="2">
                      <input name="action" type="hidden" value="thispass"<%'�����ύ%>>
&nbsp;&nbsp;
                    <input name="submit" type="submit" value="ȡ��"></td>
                  </tr>
                </table>
                <br></td>
          </tr>
        </form>
    </table></td>
  </tr>
</table>
<%end sub%>

<%sub dingdan()	'��ʾ�û�����%>
<table width="96%"  border="0" align="center" cellpadding="0" cellspacing="0">
<%
sql="select * from [dingdan] where name='"&trim(session("user"))&"';"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
if not rs.eof then
%>
  <tr>
    <td>
      <table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="BDBDBC">
          <tr>
            <td align="center"><span class="style5">�û�����</span></td>
          </tr>
          <tr>
            <td valign="top" bgcolor="#FFFFFF"><br>
                <table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="BDBDBC">
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td width="175"><div align="center">�������</div></td>
                    <td width="201"><div align="center">���ʽ</div></td>
                    <td width="215"><div align="center">�ͻ���ʽ</div></td>
                    <td width="216"><div align="center">�µ�ʱ��</div></td>
                  </tr>
				  <%do while not rs.eof%>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td><a href="chaxun.asp?dingdan=<%=rs("didanhao")%>" target="_blank"><%=rs("didanhao")%></a></td>
                    <td><%=rs("zhifu")%></td>
                    <td><%=rs("songhuo")%></td>
                    <td><%=rs("shijian")%></td>
                  </tr>
				  <%
				  rs.movenext
				  loop
				  %>
                </table>
                <br></td>
          </tr>
        </form>
    </table></td>
  </tr>
<%
end if
rs.close
set rs=nothing
%>
</table>
<%end sub%>