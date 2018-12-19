<%@LANGUAGE="VBSCRIPT"%>
<% Session.Timeout=60 %>
<!--#include file="../Conn/conn.asp" -->
<%
Sub BuildUploadRequest(RequestBin,UploadDirectory,storeType,sizeLimit,nameConflict)
  PosBeg = 1
  PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(13)))
	set checkADOConn = Server.CreateObject("ADODB.Connection")
	adoVersion = CSng(checkADOConn.Version)
	set checkADOConn = Nothing
	Length = CLng(Request.ServerVariables("HTTP_Content_Length")) 
	If "" & sizeLimit <> "" Then
    sizeLimit = CLng(sizeLimit)
    If Length > sizeLimit Then
      Request.BinaryRead (Length)
      Response.Write "<B>您上传的图片体积过大!</B><br><br>"
      Response.Write "为了版面美观，请缩小体积， <A HREF=""javascript:history.back(1)"">重新上传</a>"
      Response.End
    End If
  End If
  boundary = MidB(RequestBin,PosBeg,PosEnd-PosBeg)
  boundaryPos = InstrB(1,RequestBin,boundary)
  Do until (boundaryPos=InstrB(RequestBin,boundary & getByteString("--")))
    Dim UploadControl
    Set UploadControl = CreateObject("Scripting.Dictionary")
    Pos = InstrB(BoundaryPos,RequestBin,getByteString("Content-Disposition"))
    Pos = InstrB(Pos,RequestBin,getByteString("name="))
    PosBeg = Pos+6
    PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(34)))
    Name = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
    PosFile = InstrB(BoundaryPos,RequestBin,getByteString("filename="))
    PosBound = InstrB(PosEnd,RequestBin,boundary)
    If  PosFile<>0 AND (PosFile<PosBound) Then
      PosBeg = PosFile + 10
      PosEnd =  InstrB(PosBeg,RequestBin,getByteString(chr(34)))
      FileName = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
      FileName = Mid(FileName,InStrRev(FileName,"\")+1)
      UploadControl.Add "FileName", FileName
      Pos = InstrB(PosEnd,RequestBin,getByteString("Content-Type:"))
      PosBeg = Pos+14
      PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(13)))
      ContentType = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
      UploadControl.Add "ContentType",ContentType
      PosBeg = PosEnd+4
      PosEnd = InstrB(PosBeg,RequestBin,boundary)-2
      Value = FileName
      ValueBeg = PosBeg-1
      ValueLen = PosEnd-Posbeg
    Else
      Pos = InstrB(Pos,RequestBin,getByteString(chr(13)))
      PosBeg = Pos+4
      PosEnd = InstrB(PosBeg,RequestBin,boundary)-2
      Value = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
      ValueBeg = 0
      ValueEnd = 0
    End If
    UploadControl.Add "Value" , Value	
    UploadControl.Add "ValueBeg" , ValueBeg
    UploadControl.Add "ValueLen" , ValueLen	
    UploadRequest.Add name, UploadControl	
    BoundaryPos=InstrB(BoundaryPos+LenB(boundary),RequestBin,boundary)
  Loop

  GP_keys = UploadRequest.Keys
  for GP_i = 0 to UploadRequest.Count - 1
    GP_curKey = GP_keys(GP_i)
    '保存所有的上传文件
    if UploadRequest.Item(GP_curKey).Item("FileName") <> "" then
      GP_value = UploadRequest.Item(GP_curKey).Item("Value")
      GP_valueBeg = UploadRequest.Item(GP_curKey).Item("ValueBeg")
      GP_valueLen = UploadRequest.Item(GP_curKey).Item("ValueLen")

      if GP_valueLen = 0 then
        Response.Write "<B>发生操作错误!</B><br><br>"
        Response.Write "图片名称: " & Trim(GP_curPath) & UploadRequest.Item(GP_curKey).Item("FileName") & "<br>"
        Response.Write "该图片不存在或者是空的。<br>"
        Response.Write "请予以更正， <A HREF=""javascript:history.back(1)"">重新上传</a>"
	  	  response.End
	    end if
      
      '创建Stream对象
      Dim GP_strm1, GP_strm2
      Set GP_strm1 = Server.CreateObject("ADODB.Stream")
      Set GP_strm2 = Server.CreateObject("ADODB.Stream")
      
      '打开对象
      GP_strm1.Open
      GP_strm1.Type = 1 
      GP_strm2.Open
      GP_strm2.Type = 1 
        
      GP_strm1.Write RequestBin
      GP_strm1.Position = GP_ValueBeg
      GP_strm1.CopyTo GP_strm2,GP_ValueLen
    
      '创建并写入文件
      GP_curPath = Request.ServerVariables("PATH_INFO")
      GP_curPath = Trim(Mid(GP_curPath,1,InStrRev(GP_curPath,"/")) & UploadDirectory)
      if Mid(GP_curPath,Len(GP_curPath),1)  <> "/" then
        GP_curPath = GP_curPath & "/"
      end if 
      GP_CurFileName = UploadRequest.Item(GP_curKey).Item("FileName")
      GP_FullFileName = Trim(Server.mappath(GP_curPath))& "\" & GP_CurFileName
      GP_FileExist = false
      Set fso = CreateObject("Scripting.FileSystemObject")
      If (fso.FileExists(GP_FullFileName)) Then
        GP_FileExist = true
      End If      
      if nameConflict = "error" and GP_FileExist then
        Response.Write "<B>图片名称重复!</B><br><br>"
        Response.Write "请更改图片名称后 <A HREF=""javascript:history.back(1)"">重新上传</a>"
				GP_strm1.Close
				GP_strm2.Close
	  	  response.End
      end if
      if ((nameConflict = "over" or nameConflict = "uniq") and GP_FileExist) or (NOT GP_FileExist) then
        if nameConflict = "uniq" and GP_FileExist then
          Begin_Name_Num = 0
          while GP_FileExist    
            Begin_Name_Num = Begin_Name_Num + 1
            GP_FullFileName = Trim(Server.mappath(GP_curPath))& "\" & fso.GetBaseName(GP_CurFileName) & "_" & Begin_Name_Num & "." & fso.GetExtensionName(GP_CurFileName)
            GP_FileExist = fso.FileExists(GP_FullFileName)
          wend  
          UploadRequest.Item(GP_curKey).Item("FileName") = fso.GetBaseName(GP_CurFileName) & "_" & Begin_Name_Num & "." & fso.GetExtensionName(GP_CurFileName)
					UploadRequest.Item(GP_curKey).Item("Value") = UploadRequest.Item(GP_curKey).Item("FileName")
        end if
        on error resume next
        GP_strm2.SaveToFile GP_FullFileName,2
        if err then
          Response.Write "对不起，操作失败"
    		  err.clear
  				GP_strm1.Close
  				GP_strm2.Close
  	  	  response.End
  	    end if
  			GP_strm1.Close
  			GP_strm2.Close
  			if storeType = "path" then
  				UploadRequest.Item(GP_curKey).Item("Value") = GP_curPath & UploadRequest.Item(GP_curKey).Item("Value")
  			end if
        on error goto 0
      end if
    end if
  next
set Rs1 = Server.CreateObject("ADODB.Recordset")
sql1= "update website set banner='" & FileName & "'where id='"&Session("MR_Username")&"'"
response.Write(sql1)
rs1.open sql1,MR_website_STRING,1,3

End Sub

'把普通字符串转成二进制字符串函数
Function getByteString(StringStr)
    getByteString=""
  For i = 1 To Len(StringStr) 
    XP_varchar = mid(StringStr,i,1)
    XP_varasc = Asc(XP_varchar) 
    If XP_varasc < 0 Then 
       XP_varasc = XP_varasc + 65535 
    End If 

    If XP_varasc > 255 Then 
       XP_varlow = Left(Hex(Asc(XP_varchar)),2) 
       XP_varhigh = right(Hex(Asc(XP_varchar)),2) 
       getByteString = getByteString & chrB("&H" & XP_varlow) & chrB("&H" & XP_varhigh) 
    Else 
       getByteString = getByteString & chrB(AscB(XP_varchar)) 
    End If 
  Next 
End Function

'把二进制字符串转换成普通字符串函数 
Function getString(StringBin)
   getString =""
   Dim XP_varlen,XP_vargetstr,XP_string,XP_skip
   XP_skip = 0 
   XP_string = "" 
 If Not IsNull(StringBin) Then 
      XP_varlen = LenB(StringBin) 
    For i = 1 To XP_varlen 
      If XP_skip = 0 Then
         XP_vargetstr = MidB(StringBin,i,1) 
         If AscB(XP_vargetstr) > 127 Then 
           XP_string = XP_string & Chr(AscW(MidB(StringBin,i+1,1) & XP_vargetstr)) 
           XP_skip = 1 
         Else 
           XP_string = XP_string & Chr(AscB(XP_vargetstr)) 
         End If 
      Else 
      XP_skip = 0
   End If 
    Next 
 End If 
      getString = XP_string 
End Function 

Function UploadFormRequest(name)
  on error resume next
  if UploadRequest.Item(name) then
    UploadFormRequest = UploadRequest.Item(name).Item("Value")
  end if  
End Function

UploadQueryString = Replace(Request.QueryString,"GP_upload=true","")
if mid(UploadQueryString,1,1) = "&" then
	UploadQueryString = Mid(UploadQueryString,2)
end if

GP_uploadAction = CStr(Request.ServerVariables("URL")) & "?GP_upload=true"
If (Request.QueryString <> "") Then  
  if UploadQueryString <> "" then
	  GP_uploadAction = GP_uploadAction & "&" & UploadQueryString
  end if 
End If

If (CStr(Request.QueryString("GP_upload")) <> "") Then
  GP_redirectPage = ""
  If (GP_redirectPage = "") Then
    GP_redirectPage = CStr(Request.ServerVariables("URL"))
  end if
    
  RequestBin = Request.BinaryRead(Request.TotalBytes)
  Dim UploadRequest
  Set UploadRequest = CreateObject("Scripting.Dictionary")  
  BuildUploadRequest RequestBin, "../photos", "file", "60000", "error"
end if  
if UploadQueryString <> "" then
  UploadQueryString = UploadQueryString & "&GP_upload=true"
else  
  UploadQueryString = "GP_upload=true"
end if  

%>
<%
MR_editAction = CStr(Request.ServerVariables("URL"))
If (UploadQueryString <> "") Then
  MR_editAction = MR_editAction & "?" & UploadQueryString
End If
MR_abortEdit = false
MR_editQuery = ""
%>
<%
If (CStr(UploadFormRequest("MR_insert")) <> "") Then

  MR_editConnection = MR_website_STRING
  MR_editTable = "photos"
  MR_editRedirectUrl = "add_ok.htm"
  MR_fieldsStr  = "file0|value|id|value|other|value"
  MR_columnsStr = "tp|',none,''|id|',none,''|other|',none,''"
  MR_fields = Split(MR_fieldsStr, "|")
  MR_columns = Split(MR_columnsStr, "|")
  For i = LBound(MR_fields) To UBound(MR_fields) Step 2
    MR_fields(i+1) = CStr(UploadFormRequest(MR_fields(i)))
  Next
  If (MR_editRedirectUrl <> "" And UploadQueryString <> "") Then
    If (InStr(1, MR_editRedirectUrl, "?", vbTextCompare) = 0 And UploadQueryString <> "") Then
      MR_editRedirectUrl = MR_editRedirectUrl & "?" & UploadQueryString
    Else
      MR_editRedirectUrl = MR_editRedirectUrl & "&" & UploadQueryString
    End If
  End If

End If
%>
<%
If (CStr(UploadFormRequest("MR_insert")) <> "") Then
  MR_tableValues = ""
  MR_dbValues = ""
  For i = LBound(MR_fields) To UBound(MR_fields) Step 2
    FormVal = MR_fields(i+1)
    MR_typeArray = Split(MR_columns(i+1),",")
    Delim = MR_typeArray(0)
    If (Delim = "none") Then Delim = ""
    AltVal = MR_typeArray(1)
    If (AltVal = "none") Then AltVal = ""
    EmptyVal = MR_typeArray(2)
    If (EmptyVal = "none") Then EmptyVal = ""
    If (FormVal = "") Then
      FormVal = EmptyVal
    Else
      If (AltVal <> "") Then
        FormVal = AltVal
      ElseIf (Delim = "'") Then
        FormVal = "'" & Replace(FormVal,"'","''") & "'"
      Else
        FormVal = Delim + FormVal + Delim
      End If
    End If
    If (i <> LBound(MR_fields)) Then
      MR_tableValues = MR_tableValues & ","
      MR_dbValues = MR_dbValues & ","
    End if
    MR_tableValues = MR_tableValues & MR_columns(i)
    MR_dbValues = MR_dbValues & FormVal
  Next
  MR_editQuery = "insert into " & MR_editTable & "(" & MR_tableValues & ") values (" & MR_dbValues & ")"

  If (Not MR_abortEdit) Then
    '向数据表中写入数据
    Set MR_editCmd = Server.CreateObject("ADODB.CoMMand")
    MR_editCmd.ActiveConnection = MR_editConnection
    MR_editCmd.CoMMandText = MR_editQuery
    MR_editCmd.Execute
    MR_editCmd.ActiveConnection.Close

    If (MR_editRedirectUrl <> "") Then
      Response.Redirect(MR_editRedirectUrl)
    End If
  End If

End If
%>
<%
Dim ps__MRColParam
ps__MRColParam = "1"
if (Session("MR_username") <> "") then ps__MRColParam = Session("MR_username")
%>
<%
set ps = Server.CreateObject("ADODB.Recordset")
ps.ActiveConnection = MR_website_STRING
ps.Source = "SELECT *  FROM photos WHERE id = '" + Replace(ps__MRColParam, "'", "''") + "' and other = 'ok'  ORDER BY idd DESC"
ps.CursorType = 0
ps.CursorLocation = 2
ps.LockType = 3
ps.Open()
ps_numRows = 0
%>
<%
Dim Rs__MRColParam
Rs__MRColParam = "1"
if (Session("MR_username") <> "") then Rs__MRColParam = Session("MR_username")
%>
<%
set Rs = Server.CreateObject("ADODB.Recordset")
Rs.ActiveConnection = MR_website_STRING
Rs.Source = "SELECT * FROM website WHERE id = '" + Replace(Rs__MRColParam, "'", "''") + "'"
Rs.CursorType = 0
Rs.CursorLocation = 2
Rs.LockType = 3
Rs.Open()
Rs_numRows = 0

%>
<%
Dim Repeat1__numRows
Repeat1__numRows = 10
Dim Repeat1__index
Repeat1__index = 0
ps_numRows = ps_numRows + Repeat1__numRows
%>
<%
ps_total = ps.RecordCount
If (ps_numRows < 0) Then
  ps_numRows = ps_total
Elseif (ps_numRows = 0) Then
  ps_numRows = 1
End If
ps_first = 1
ps_last  = ps_first + ps_numRows - 1
If (ps_total <> -1) Then
  If (ps_first > ps_total) Then ps_first = ps_total
  If (ps_last > ps_total) Then ps_last = ps_total
  If (ps_numRows > ps_total) Then ps_numRows = ps_total
End If
%>
<%
If (ps_total = -1) Then
  ps_total=0
  While (Not ps.EOF)
    ps_total = ps_total + 1
    ps.MoveNext
  Wend
  If (ps.CursorType > 0) Then
    ps.MoveFirst
  Else
    ps.Requery
  End If
  If (ps_numRows < 0 Or ps_numRows > ps_total) Then
    ps_numRows = ps_total
  End If
  ps_first = 1
  ps_last = ps_first + ps_numRows - 1
  If (ps_first > ps_total) Then ps_first = ps_total
  If (ps_last > ps_total) Then ps_last = ps_total

End If
%>
<%
Set MR_rs    = ps
MR_rsCount   = ps_total
MR_size      = ps_numRows
MR_uniqueCol = ""
MR_paramName = ""
MR_offset = 0
MR_atTotal = false
MR_paramIsDefined = false
If (MR_paramName <> "") Then
  MR_paramIsDefined = (Request.QueryString(MR_paramName) <> "")
End If
%>
<%
if (Not MR_paramIsDefined And MR_rsCount <> 0) then
  r = Request.QueryString("index")
  If r = "" Then r = Request.QueryString("offset")
  If r <> "" Then MR_offset = Int(r)
  If (MR_rsCount <> -1) Then
    If (MR_offset >= MR_rsCount Or MR_offset = -1) Then  
      If ((MR_rsCount Mod MR_size) > 0) Then  
        MR_offset = MR_rsCount - (MR_rsCount Mod MR_size)
      Else
        MR_offset = MR_rsCount - MR_size
      End If
    End If
  End If
  i = 0
  While ((Not MR_rs.EOF) And (i < MR_offset Or MR_offset = -1))
    MR_rs.MoveNext
    i = i + 1
  Wend
  If (MR_rs.EOF) Then MR_offset = i 

End If
%>
<%
If (MR_rsCount = -1) Then
  i = MR_offset
  While (Not MR_rs.EOF And (MR_size < 0 Or i < MR_offset + MR_size))
    MR_rs.MoveNext
    i = i + 1
  Wend
  If (MR_rs.EOF) Then
    MR_rsCount = i
    If (MR_size < 0 Or MR_size > MR_rsCount) Then MR_size = MR_rsCount
  End If
  If (MR_rs.EOF And Not MR_paramIsDefined) Then
    If (MR_offset > MR_rsCount - MR_size Or MR_offset = -1) Then
      If ((MR_rsCount Mod MR_size) > 0) Then
        MR_offset = MR_rsCount - (MR_rsCount Mod MR_size)
      Else
        MR_offset = MR_rsCount - MR_size
      End If
    End If
  End If
  If (MR_rs.CursorType > 0) Then
    MR_rs.MoveFirst
  Else
    MR_rs.Requery
  End If
  i = 0
  While (Not MR_rs.EOF And i < MR_offset)
    MR_rs.MoveNext
    i = i + 1
  Wend
End If
%>
<%
ps_first = MR_offset + 1
ps_last  = MR_offset + MR_size
If (MR_rsCount <> -1) Then
  If (ps_first > MR_rsCount) Then ps_first = MR_rsCount
  If (ps_last > MR_rsCount) Then ps_last = MR_rsCount
End If
MR_atTotal = (MR_rsCount <> -1 And MR_offset + MR_size >= MR_rsCount)
%>
<%
MR_removeList = "&index="
If (MR_paramName <> "") Then MR_removeList = MR_removeList & "&" & MR_paramName & "="
MR_keepURL="":MR_keepForm="":MR_keepBoth="":MR_keepNone=""
For Each Item In Request.QueryString
  NextItem = "&" & Item & "="
  If (InStr(1,MR_removeList,NextItem,1) = 0) Then
    MR_keepURL = MR_keepURL & NextItem & Server.URLencode(Request.QueryString(Item))
  End If
Next
For Each Item In Request.Form
  NextItem = "&" & Item & "="
  If (InStr(1,MR_removeList,NextItem,1) = 0) Then
    MR_keepForm = MR_keepForm & NextItem & Server.URLencode(Request.Form(Item))
  End If
Next
MR_keepBoth = MR_keepURL & MR_keepForm
if (MR_keepBoth <> "") Then MR_keepBoth = Right(MR_keepBoth, Len(MR_keepBoth) - 1)
if (MR_keepURL <> "")  Then MR_keepURL  = Right(MR_keepURL, Len(MR_keepURL) - 1)
if (MR_keepForm <> "") Then MR_keepForm = Right(MR_keepForm, Len(MR_keepForm) - 1)
Function MR_joinChar(firstItem)
  If (firstItem <> "") Then
    MR_joinChar = "&"
  Else
    MR_joinChar = ""
  End If
End Function
%>
<%
MR_keepMove = MR_keepBoth
MR_moveParam = "index"
If (MR_size > 0) Then
  MR_moveParam = "offset"
  If (MR_keepMove <> "") Then
    params = Split(MR_keepMove, "&")
    MR_keepMove = ""
    For i = 0 To UBound(params)
      nextItem = Left(params(i), InStr(params(i),"=") - 1)
      If (StrComp(nextItem,MR_moveParam,1) <> 0) Then
        MR_keepMove = MR_keepMove & "&" & params(i)
      End If
    Next
    If (MR_keepMove <> "") Then
      MR_keepMove = Right(MR_keepMove, Len(MR_keepMove) - 1)
    End If
  End If
End If
If (MR_keepMove <> "") Then MR_keepMove = MR_keepMove & "&"
urlStr = Request.ServerVariables("URL") & "?" & MR_keepMove & MR_moveParam & "="
MR_moveFirst = urlStr & "0"
MR_moveLast  = urlStr & "-1"
MR_moveNext  = urlStr & Cstr(MR_offset + MR_size)
prev = MR_offset - MR_size
If (prev < 0) Then prev = 0
MR_movePrev  = urlStr & Cstr(prev)
%>
<%
Dim Hi_Repeat__numRows_Cu,Hi_Pages1_Cu,Hi_Pages2_Cu
    Hi_Repeat__numRows_Cu = 10
%>
<%
Dim Hi_Repeat__numRows,Hi_Pages1,Hi_Pages2
    Hi_Repeat__numRows = 10
%>
<SCRIPT RUNAT=SERVER LANGUAGE=VBSCRIPT>					
function DoDateTime(str, nNamedFormat, nLCID)				
	dim strRet								
	dim nOldLCID								
										
	strRet = str								
	If (nLCID > -1) Then							
		oldLCID = Session.LCID						
	End If									
										
	On Error Resume Next							
										
	If (nLCID > -1) Then							
		Session.LCID = nLCID						
	End If									
										
	If ((nLCID < 0) Or (Session.LCID = nLCID)) Then				
		strRet = FormatDateTime(str, nNamedFormat)			
	End If									
										
	If (nLCID > -1) Then							
		Session.LCID = oldLCID						
	End If									
										
	DoDateTime = strRet							
End Function									
</SCRIPT>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>在线图片上传</title>
<style type="text/css">
<!--
.pt9 {  font-family: "宋体"; font-size: 9pt}
.pt9n {  font-family: "宋体"; font-size: 9pt; line-height: 13pt}
a:hover {  color: #0000FF; text-decoration: underline}
a:link {  color: #0000FF}
a:visited {  text-decoration: none}
-->
</style>
<script language="JavaScript">
<!--

function getFileExtension(filePath) { 
  fileName = ((filePath.indexOf('/') > -1) ? filePath.substring(filePath.lastIndexOf('/')+1,filePath.length) : filePath.substring(filePath.lastIndexOf('\\')+1,filePath.length));
  return fileName.substring(fileName.lastIndexOf('.')+1,fileName.length);
}

function checkFileUpload(form,extensions) { //v1.0
  document.MR_returnValue = true;
  if (extensions && extensions != '') {
    for (var i = 0; i<form.elements.length; i++) {
      field = form.elements[i];
      if (field.type.toUpperCase() != 'FILE') continue;
      if (field.value == '') {
        alert('文件框中必须保证已经有文件被选中!');
        document.MR_returnValue = false;field.focus();break;
      }
      if (extensions.toUpperCase().indexOf(getFileExtension(field.value).toUpperCase()) == -1) {
        alert('这种文件类型不允许上传!.\n只有以下类型的文件才被允许上传: ' + extensions + '.\n请选择别的文件并重新上传.');
        document.MR_returnValue = false;field.focus();break;
  } } }
}
//-->
</script>
</head>
<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF" marginwidth="0" marginheight="0">
<form name="form1" method="POST" action="<%=MR_editAction%>" enctype="multipart/form-data" onSubmit="checkFileUpload(this,'GIF,JPG,JPEG');return document.MR_returnValue">
  <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#B9B9B9" class="pt9">
    <tr bgcolor="#FF9900">
      <td height="24" colspan="2" bgcolor="#33CCFF" class="s9back">&nbsp;<font color="#FFFFFF">如果您需要的图片还没上传,您可以在此直接上传.</font></td>
    </tr>
    <tr bgcolor="#FF9900">
      <td width="20%" height="32" bgcolor="#33CCFF" class="s9back"><div align="center"><b><font color="#FFFFFF">上传图片：</font></b></div></td>
      <td height="32" bgcolor="#33CCFF" class="s9back"><input type="file"  name="file0" class="pt9">
      </td>
    </tr>
    <tr bgcolor="#FF9900">
      <td height="32" colspan="2" align="center" bgcolor="#33CCFF" class="s9back"><input type="submit" name="Submit22"  style="background-color=#b6c5d8"  value="添加图片" class="pt9">
        <input type="hidden" name="id" value="<%= Session("MR_username") %>">
      <input type="hidden" name="other" value="ok">      </td>
    </tr>
  </table>
  <input type="hidden" name="MR_insert" value="true">
</form>
</body>
</html>
<%
ps.Close()
%>
