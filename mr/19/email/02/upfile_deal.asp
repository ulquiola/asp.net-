<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="upload_wj.inc"-->
<%
set upload=new upload_myfile
if upload.form("act")="uploadfile" then
  filepath=trim(upload.form("filepath"))  '�ϴ��ļ���ָ����·��
  filelx=trim(upload.form("filelx"))  '�ϴ��ļ�����,����Ĭ��Ϊjpg	
  i=0
  for each file_name in upload.File
   set file=upload.File(file_name) 
   fileExt=lcase(file.FileExt)	
   if file.filesize<100 then
	 response.write "<script language=javascript>alert('����ѡ���ļ���');history.go(-1);</script>"
	 response.end
   end if 
  if filelx="jpg" then
	if file.filesize>(1000*1024) then
	  response.write "<script language=javascript>alert('ͼƬ�ļ���С���ܳ���1M��');history.go(-1);</script>"
	  response.end
	end if
  end if
  '�����������
  randomize
  ranNum=int(90000*rnd)+10000
  filename=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&"."&fileExt
  
  if file.FileSize>0 then         
    file.SaveToFile Server.mappath(FileName)
    if filelx="swf" then
      response.write "<script>window.opener.document."&upload.form("FormName")&".size.value='"&int(file.FileSize/1024)&" K'</script>"
    end if
    response.write "<script>window.opener.document."&upload.form("FormName")&"."&upload.form("EditName")&".value='"&FileName&"'</script>"
  end if
  set file=nothing
  next
  set upload=nothing
end if
%>
<script language="javascript">
window.alert("�ļ��ϴ��ɹ�!�벻Ҫ�޸����ɵ�����·����");
window.close();
</script>