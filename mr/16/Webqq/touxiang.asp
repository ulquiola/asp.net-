<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/conn.asp" -->
<!--#include file="js/Function.asp" -->
       <%
        CurrUserID =  ",S" & Request.Cookies("UserID") & "E"
	    application("OnLineUserID") = replace(application("OnLineUserID"),CurrUserID,"")
	    application("OnLineUserID") = replace(application("OnLineUserID"),",SE","")
	    application("OnLineUserID") = application("OnLineUserID") & CurrUserID
	    
		Set TalkHistory = Server.CreateObject("ADODB.Recordset")
		TalkHistory.ActiveConnection = conn
		TalkHistory.Source = "SELECT * FROM tb_talk WHERE  ToUserID = " &request.Cookies("UserID")&" and  Isreaded = '0' order by id desc"
		TalkHistory.CursorType = 0
		TalkHistory.CursorLocation = 2
		TalkHistory.LockType = 1
		TalkHistory.Open()
		%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
<script>
var OnLineUserID  = "<%=application("OnLineUserID")%>"
//alert(OnLineUserID)
OnLineUserIDAry = OnLineUserID.split(",")
if(window.parent.document.all.faceImg){
	if(window.parent.document.all.faceImg.length){
	   for(i=0;i<window.parent.document.all.faceImg.length;i++){
	    UserID = "S"+window.parent.document.all.faceImg[i].id.replace("Img","")+"E"
	    
			var IsOnline = "0"
			for(j=0;j<OnLineUserIDAry.length;j++){
			    if(OnLineUserIDAry[j]==UserID){
			       IsOnline = "1"
			    }
	   		}	
	   		if(IsOnline=="1"){
				var tempStr=window.parent.document.all.faceImg[i].src;
				if(tempStr.lastIndexOf("_0")>1){	//控制上线时显示的图片
					tempStr=tempStr.substring(0,tempStr.lastIndexOf("_0"))+".gif";
				}
			 	window.parent.document.all.faceImg[i].src=tempStr
	   		}
	   		else{
				var tempStr=window.parent.document.all.faceImg[i].src;

				if(tempStr.lastIndexOf("_0")<1){//控制下线时显示的图片
								//alert(tempStr.lastIndexOf("_0"));
					tempStr=tempStr.substring(0,tempStr.lastIndexOf("."))+"_0.gif";
				}
				window.parent.document.all.faceImg[i].src=tempStr;
	   		}
	   }
	}else{
		//alert(window.parent.document.all.faceImg.id);
	 	UserID = "S"+window.parent.document.all.faceImg.id.replace("Img","")+"E"
		 var IsOnline = "0"
		for(j=0;j<OnLineUserIDAry.length;j++){
		    if(OnLineUserIDAry[j]==UserID){
		       IsOnline = "1"
		    }
	   	}	
	   	if(IsOnline=="1"){
			var tempStr=window.parent.document.all.faceImg.src;
			if(tempStr.lastIndexOf("_0")!=-1){		//控制上线时显示的图片
				tempStr=tempStr.substring(0,tempStr.lastIndexOf("_"))+".gif";
			}
			window.parent.document.all.faceImg.src=tempStr
	   	}else{
			var tempStr=window.parent.document.all.faceImg.src;
			
			if(tempStr.lastIndexOf("_0")<1){		//控制下线时显示的图片
				tempStr=tempStr.substring(0,tempStr.lastIndexOf("."))+"_0.gif";
			}

			window.parent.document.all.faceImg.src=tempStr
	   	}
	}
}
<%
while not TalkHistory.EOF
'主要用于判断用户是否接受到信息
%>
if(window.parent.document.all.Img<%=talkHistory("fuserid")%>){
			var tempStr1=window.parent.document.all.Img<%=talkHistory("fuserid")%>.src;
			if(tempStr1.lastIndexOf("_")!=-1){
				tempStr1=tempStr1.substring(0,tempStr1.lastIndexOf("_"))+"_00.gif";
			}else{
				tempStr1=tempStr1.substring(0,tempStr1.lastIndexOf("."))+"_00.gif";
			}
			window.parent.document.all.Img<%=talkHistory("fuserid")%>.src=tempStr1;
Url = "MsgAlert.htm"
MsgAlert = window.open(Url,"MsgAlert","width=150,height=100")
}
else{
window.parent.document.all.unknowMsgTD.HasNewMsg="1"
window.parent.document.all.unknowMsgTD.UserID = "<%=talkHistory("fuserid")%>"
window.parent.document.all.unknowMsg.style.display = ""
var tempStr1=window.parent.document.all.Img<%=talkHistory("fuserid")%>.src;
tempStr1=tempStr1.substring(0,tempStr1.lastIndexOf("_"))+".gif";
window.parent.document.all.Img<%=talkHistory("fuserid")%>.src=tempStr1;
Url = "MsgAlert.htm"
MsgAlert = window.open(Url,"MsgAlert","width=150,height=100")
}
<%
TalkHistory.moveNext()
wend
talkHistory.close()
set TalkHistory = nothing
%>
</script>
</head>
<body>
</body>
</html>
