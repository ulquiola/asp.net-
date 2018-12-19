<%
yanzheng(request.QueryString("num"))
Function yanzheng(num)
	Dim codeimage(4),NumStr
		NumStr = num
		For i=0 To 3
			codeimage(i)=mid(NumStr,i+1,1)
		Next
	Set obj=Server.CreateObject("Adodb.Stream")
		obj.Mode=3
		obj.Type=1
		obj.Open
	Set obj1=Server.CreateObject("Adodb.Stream")
		obj1.Mode=3
		obj1.Type=1
		obj1.Open
		obj.LoadFromFile(Server.mappath("sun2.Fix"))
		obj1.write obj.read(1280)
		For i=0 To 3
			obj.Position=(9-codeimage(i))*320
			obj1.Position=i*320
			obj1.write obj.read(320)
		Next 	
		obj.LoadFromFile(Server.mappath("sun1.fix"))
		pp=lenb(obj.read())
		obj.Position=pp
		For i=0 To 9 Step 1
			For j=0 To 3
				obj1.Position=i*32+j*320
				obj.Position=pp+30*j+i*120
				obj.write obj1.read(30)
			Next 
		Next 
		Response.ContentType = "image/BMP"
		obj.Position=0
		Response.BinaryWrite obj.read()
End Function
%>