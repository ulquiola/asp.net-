<%
dim ouf

Class upload_myfile
  
dim Form,File,Version
  
Private Sub Class_Initialize 
  dim date,start,bCrLf,startInfo,iInfoStart,Endinfo,tstr,iStart,fif
  dim ifs,ifp,sft,sfva,sfn
  dim ifds,ifed
  dim ifst,ifoe,sfna

  set Form = Server.CreateObject("Scripting.Dictionary")
  set File = Server.CreateObject("Scripting.Dictionary")
  if Request.TotalBytes < 1 then Exit Sub
  set tstr = Server.CreateObject("adodb.stream")
  set ouf = Server.CreateObject("adodb.stream")
  ouf.Type = 1
  ouf.Mode = 3
  ouf.Open 
  ouf.Write Request.BinaryRead(Request.TotalBytes)
  ouf.Position=0
  date = ouf.Read 
  ifoe = ouf.Size
  bCrLf = chrB(13) & chrB(10)
  start = MidB(date,1, InStrB(1,date,bCrLf)-1)
  iStart = LenB (start)
  ifst = iStart+2
  Do
    Endinfo = InStrB(ifst,date,bCrLf & bCrLf)+3
    tstr.Type = 1
    tstr.Mode = 3
    tstr.Open
    ouf.Position = ifst
    ouf.CopyTo tstr,Endinfo-ifst
    tstr.Position = 0
    tstr.Type = 2
    tstr.Charset ="gb2312"
    startInfo = tstr.ReadText      
    ifst = InStrB(Endinfo,date,start)-1
    ifds = InStr(22,startInfo,"name=""",1)+6
    ifed = InStr(ifds,startInfo,"""",1)
    sfna = Mid (startInfo,ifds,ifed-ifds)
    if InStr (45,startInfo,"filename=""",1) > 0 then
      set fif= new FileInfo
      ifds = InStr(ifed,startInfo,"filename=""",1)+10
      ifed = InStr(ifds,startInfo,"""",1)
      sfn = Mid (startInfo,ifds,ifed-ifds)
      '获得文件路径
      If sfn <> "" Then
        fif.FilePath = left(sfn,InStrRev(sfn, "\"))
      Else
        fif.FilePath = ""
      End If
	  '获得文件名
	  If sfn <> "" Then
        fif.FileName = mid(sfn,InStrRev(sfn, "\")+1)
      Else
        fif.FileName = ""
      End If
      '获得文件扩展名
	  If sfn <> "" Then
		 fif.FileExt = mid(sfn,InStrRev(sfn, ".")+1)
	  Else
		 fif.FileExt = ""
	  End If
      ifds = InStr(ifed,startInfo,"Content-Type: ",1)+14
      ifed = InStr(ifds,startInfo,vbCr)
      fif.FileType = Mid (startInfo,ifds,ifed-ifds)
      fif.FileStart = Endinfo
      fif.FileSize = ifst -Endinfo -2
      fif.FormName = sfna
      file.add sfna,fif
    else
      tstr.Close
      tstr.Type = 1
      tstr.Mode = 3
      tstr.Open
      ouf.Position = Endinfo 
      ouf.CopyTo tstr,ifst-Endinfo-2
      tstr.Position = 0
      tstr.Type = 2
      tstr.Charset = "gb2312"
      sfva = tstr.ReadText 
      form.Add sfna,sfva
    end if
    tstr.Close
    ifst = ifst+iStart+2
    loop until (ifst+2) = ifoe 
  date=""
  set tstr = nothing
End Sub

Private Sub Class_Terminate  
  if not Request.TotalBytes<1 then
    ouf.Close
    set ouf =nothing
    end if
  Form.RemoveAll
  File.RemoveAll
  set Form=nothing
  set File=nothing
End Sub
   
End Class

Class FileInfo
  dim FormName,FileName,FilePath,FileSize,FileType,FileStart,FileExt
  Private Sub Class_Initialize 
    FileName = ""
    FilePath = ""
    FileSize = 0
    FileStart= 0
    FormName = ""
    FileType = ""
    FileExt = ""
  End Sub
  
 Public function SaveToFile(FullPath)
    dim oFileStream,ErrorChar,i
    SaveToFile=1
    if trim(fullpath)="" or right(fullpath,1)="/" then exit function
    set oFileStream=CreateObject("Adodb.Stream")
    oFileStream.Type=1
    oFileStream.Mode=3
    oFileStream.Open
    ouf.position=FileStart
    ouf.copyto oFileStream,FileSize

    oFileStream.SaveToFile FullPath,2
    oFileStream.Close
    set oFileStream=nothing 
    SaveToFile=0
  end function
End Class
%>