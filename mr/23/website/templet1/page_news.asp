<!--#include file="../Conn/conn.asp" -->
<%
if(Request.QueryString("id") <> "") then hit__MRColParam = Request.QueryString("id")
%>
<%
Dim p1__MRColParam
p1__MRColParam = "1"
if (Request.QueryString("id") <> "") then p1__MRColParam = Request.QueryString("id")
%>
<%
set p1 = Server.CreateObject("ADODB.Recordset")
p1.ActiveConnection = MR_website_STRING
p1.Source = "SELECT * FROM website WHERE id = '" + Replace(p1__MRColParam, "'", "''") + "'"
p1.CursorType = 0
p1.CursorLocation = 2
p1.LockType = 3
p1.Open()
p1_numRows = 0
%>
<%
set hit = Server.CreateObject("ADODB.CoMMand")
hit.ActiveConnection = MR_website_STRING
hit.CoMMandText = "UPDATE website  SET hit =hit+1  WHERE id ='" + Replace(hit__MRColParam, "'", "''") + "'"
hit.CoMMandType = 1
hit.CoMMandTimeout = 0
hit.Prepared = true
hit.Execute()
%>
<%
Dim news__MRColParam
news__MRColParam = "1"
if (Request.QueryString("id") <> "") then news__MRColParam = Request.QueryString("id")
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = 20
Dim Repeat1__index
Repeat1__index = 0
news_numRows = news_numRows + Repeat1__numRows
%>
<%
if (Not MR_paramIsDefined And MR_rsCount <> 0) then
  r = Request.QueryString("index")
  If r = "" Then r = Request.QueryString("offset")
  If r <> "" Then MR_offset = Int(r)
  If (MR_rsCount <> -1) Then
    If (MR_offset >= MR_rsCount Or MR_offset = -1) Then  ' past end or move last
      If ((MR_rsCount Mod MR_size) > 0) Then         ' last page not a full repeat region
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
news_first = MR_offset + 1
news_last  = MR_offset + MR_size
If (MR_rsCount <> -1) Then
  If (news_first > MR_rsCount) Then news_first = MR_rsCount
  If (news_last > MR_rsCount) Then news_last = MR_rsCount
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
    Hi_Repeat__numRows_Cu = 20
%>
<%
Dim Hi_Repeat__numRows,Hi_Pages1,Hi_Pages2
    Hi_Repeat__numRows = 20
%>
<html>
<head>
<title><%=(p1.Fields.Item("title").Value)%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link 
href="../website/templet1/css/css.css" rel=stylesheet type="text/css">
<style type="text/css">
<!--
.pt9 {  font-family: "宋体"; font-size: 9pt; line-height: 14pt}
-->
</style>
</head>
<% If (p1.Fields.Item("floatlogo").Value) <> ("00.bmp") Then 'script %>
<script LANGUAGE="JavaScript">
var brOK=false;
var mie=false;
var aver=parseInt(navigator.appVersion.substring(0,1));
var aname=navigator.appName;
function checkbrOK()
{if(aname.indexOf("Internet Explorer")!=-1)
   {if(aver>=4) brOK=navigator.javaEnabled();
    mie=true;
   }
 if(aname.indexOf("Netscape")!=-1)
   {if(aver>=4) brOK=navigator.javaEnabled();}
}
var vmin=2;
var vmax=5;
var vr=0.02;
var timer1;
function Chip(chipname,width,height)
{this.named=chipname;
 this.vx=vmin+vmax*Math.random();
 this.vy=vmin+vmax*Math.random();
 this.w=width;
 this.h=height;
 this.xx=0;
 this.yy=0;
 this.timer1=null;
}
function movechip(chipname)
{
 if(brOK)
  {eval("chip="+chipname);
   if(!mie)
    {pageX=window.pageXOffset;
     pageW=window.innerWidth;
     pageY=window.pageYOffset;
     pageH=window.innerHeight;
    }
   else
    {pageX=window.document.body.scrollLeft;
     pageW=window.document.body.offsetWidth-22;
     pageY=window.document.body.scrollTop;
     pageH=window.document.body.offsetHeight-34;
    }

   chip.xx=chip.xx+chip.vx;
   chip.yy=chip.yy+chip.vy;

   chip.vx+=vr*(Math.random()-0.5);
   chip.vy+=vr*(Math.random()-0.5);
   if(chip.vx>(vmax+vmin))  chip.vx=(vmax+vmin)*2-chip.vx;
   if(chip.vx<(-vmax-vmin)) chip.vx=(-vmax-vmin)*2-chip.vx;
   if(chip.vy>(vmax+vmin))  chip.vy=(vmax+vmin)*2-chip.vy;
   if(chip.vy<(-vmax-vmin)) chip.vy=(-vmax-vmin)*2-chip.vy;


   if(chip.xx<=pageX)
     {chip.xx=pageX;
      chip.vx=vmin+vmax*Math.random();
     }
   if(chip.xx>=pageX+pageW-chip.w)
     {chip.xx=pageX+pageW-chip.w;
      chip.vx=-vmin-vmax*Math.random();
     }
   if(chip.yy<=pageY)
     {chip.yy=pageY;
      chip.vy=vmin+vmax*Math.random();
     }
   if(chip.yy>=pageY+pageH-chip.h)
     {chip.yy=pageY+pageH-chip.h;
      chip.vy=-vmin-vmax*Math.random();
     }

   if(!mie)
      {eval('document.'+chip.named+'.top ='+chip.yy);
       eval('document.'+chip.named+'.left='+chip.xx);
      }
   else
      {eval('document.all.'+chip.named+'.style.pixelLeft='+chip.xx);
       eval('document.all.'+chip.named+'.style.pixelTop ='+chip.yy);
      }
   chip.timer1=setTimeout("movechip('"+chip.named+"')",140);
  }
}

function hide(chipname){
        if(brOK){
                eval("chip="+chipname);
                if(!mie)
                        eval('document.'+chip.named+'.visibility ='+"'hide'");
                else
                        eval('document.all.'+chip.named+'.style.visibility ='+"'hidden'");
        }
}

function stopme(chipname)
{if(brOK)
  {//alert(chipname)
   eval("chip="+chipname);
   if(chip.timer1!=null)
    {clearTimeout(chip.timer1)}
  }
}
var chip;
function pagestart()
{checkbrOK();
 chip=new Chip("chip",117,75);
  if(brOK)
   { movechip("chip");
    }
}
</script>
<% End If%>
<body bgcolor="#EFEFEF" text="#000000" leftmargin="0" topmargin="0">
<table width="778" border="0" cellspacing="0" cellpadding="0" height="79" align="center" background="images/1-1.gif">
  <tr>
    <td height="61"><% If (p1.Fields.Item("logo").Value) <> ("0.bmp") Then %>
      <img src="../photos/<%=(p1.Fields.Item("logo").Value)%>" width="<%=(p1.Fields.Item("logo_w").Value)%>" height="<%=(p1.Fields.Item("logo_h").Value)%>">
      <% End If %>
    </td>
    <td height="61"><% If p1.Fields.Item("banner").Value = ("000.bmp") Then %>
      <div align="right"; style="font size=<%=(p1.Fields.Item("titlesize").Value)%>; FILTER: <%=(p1.Fields.Item("cssopen").Value)%>(color=<%=(p1.Fields.Item("titlecss").Value)%>,direction=<%=(p1.Fields.Item("titleway").Value)%>); WIDTH: 100%; COLOR: <%=(p1.Fields.Item("titlecolor").Value)%>; FONT-FAMILY: <%=(p1.Fields.Item("titlefont").Value)%>"><b><%=(p1.Fields.Item("title").Value)%>&nbsp;</b></div>
      <% End If %>
      <% If p1.Fields.Item("banner").Value <> ("000.bmp") Then %>
      <div align="right"><img src="../photo/<%=(p1.Fields.Item("banner").Value)%>" width="<%=(p1.Fields.Item("banner_w").Value)%>" height="<%=(p1.Fields.Item("banner_h").Value)%>"></div>
      <% End If %></td>
  </tr>
</table>
<table width="778" border="0" cellspacing="0" cellpadding="0" align="center" height="29">
  <tr>
    <td height="22" bgcolor="#FFFFFF"><table width="70%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
          <td height="22" bgcolor="#FFFFFF"><marquee>
            <font color="#000000">
            <% If (p1.Fields.Item("yin_zimu").Value) <> ("Y") Then %>
            <span class="pt9"><%=(p1.Fields.Item("zimu").Value)%></span></font> <font color="#000000">
            <% End If %>
            </font>
            </marquee></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td height="1" bgcolor="#FFFFFF"></td>
  </tr>
  <tr>
    <td height="8" bgcolor="#999999"></td>
  </tr>
  <tr>
    <td height="27" bgcolor="#DFF0F7"><table width="85%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="27"><div align="right">
              <script language="JavaScript">
 today=new Date();
 function initArray(){
   this.length=initArray.arguments.length
   for(var i=0;i<this.length;i++)
   this[i+1]=initArray.arguments[i]  }
   var d=new initArray(
     "星期日",
     "星期一",
     "星期二",
     "星期三",
     "星期四",
     "星期五",
     "星期六");
document.write(
     "<font color=#000000 style='font-size:9pt;font-family: 宋体'> ",
     today.getYear(),"年",
     today.getMonth()+1,"月",
     today.getDate(),"日&nbsp;",
     d[today.getDay()+1],
     "</font>" );
</script>
            </div></td>
        </tr>
    </table></td>
  </tr>
</table>
<table width="778" border="0" cellspacing="0" cellpadding="0" align="center" height="379">
  <tr>
    <td width="156" valign="top" height="356" bgcolor="#DFF0F7"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="25" bgcolor="#FFFFFF"><div align="center"><img src="images/5-4.gif" width="79" height="19"></div></td>
        </tr>
        <% If (p1.Fields.Item("yin_1").Value) <> ("Y") Then 'script %>
        <tr>
          <td height="25"><table width="100%" border="0" cellspacing="0" cellpadding="0" height="25">
              <tr>
                <td width="22%">&nbsp;</td>
                <td><a href="page1.asp?<%= MR_keepNone & MR_joinChar(MR_keepNone) & "id=" & p1.Fields.Item("id").Value %>" class="<%=(p1.Fields.Item("font1").Value)%>"><%=(p1.Fields.Item("title1").Value)%></a> </td>
              </tr>
            </table></td>
        </tr>
        <% End If %>
        <% If (p1.Fields.Item("yin_news").Value) <> ("Y") Then %>
        <tr>
          <td height="25"><table width="100%" border="0" cellspacing="0" cellpadding="0" height="25">
              <tr>
                <td width="22%">&nbsp;</td>
                <td><a href="page_news.asp?<%= MR_keepNone & MR_joinChar(MR_keepNone) & "id=" & p1.Fields.Item("id").Value %>" class="<%=(p1.Fields.Item("font1").Value)%>"><%=(p1.Fields.Item("title_news").Value)%></a> </td>
              </tr>
            </table></td>
        </tr>
        <% End If %>
        <% If (p1.Fields.Item("yin_3").Value) <> ("Y") Then %>
        <tr>
          <td height="25"><table width="100%" border="0" cellspacing="0" cellpadding="0" height="25">
              <tr>
                <td width="22%">&nbsp;</td>
                <td><a href="page3.asp?<%= MR_keepNone & MR_joinChar(MR_keepNone) & "id=" & p1.Fields.Item("id").Value %>" class="<%=(p1.Fields.Item("font1").Value)%>"><%=(p1.Fields.Item("title3").Value)%></a> </td>
              </tr>
            </table></td>
        </tr>
        <% End If %>
        <% If (p1.Fields.Item("yin_4").Value) <> ("Y") Then %>
        <tr>
          <td height="25"><table width="99%" border="0" cellspacing="0" cellpadding="0" height="25">
              <tr>
                <td width="22%">&nbsp;</td>
                <td><a href="page4.asp?<%= MR_keepNone & MR_joinChar(MR_keepNone) & "id=" & p1.Fields.Item("id").Value %>" class="<%=(p1.Fields.Item("font1").Value)%>"><%=(p1.Fields.Item("title4").Value)%></a> </td>
              </tr>
            </table></td>
        </tr>
        <% End If%>
        <% If (p1.Fields.Item("yin_5").Value) <> ("Y") Then %>
        <tr>
          <td height="25"><table width="100%" border="0" cellspacing="0" cellpadding="0" height="25">
              <tr>
                <td width="22%">&nbsp;</td>
                <td><a href="page5.asp?<%= MR_keepNone & MR_joinChar(MR_keepNone) & "id=" & p1.Fields.Item("id").Value %>" class="<%=(p1.Fields.Item("font1").Value)%>"><%=(p1.Fields.Item("title5").Value)%></a> </td>
              </tr>
            </table></td>
        </tr>
        <% End If %>
        <% If (p1.Fields.Item("yin_6").Value) <> ("Y") Then %>
        <tr>
          <td height="25"><table width="100%" border="0" cellspacing="0" cellpadding="0" height="25">
              <tr>
                <td width="22%">&nbsp;</td>
                <td><a href="page6.asp?<%= MR_keepNone & MR_joinChar(MR_keepNone) & "id=" & p1.Fields.Item("id").Value %>" class="<%=(p1.Fields.Item("font1").Value)%>"><%=(p1.Fields.Item("title6").Value)%></a> </td>
              </tr>
            </table></td>
        </tr>
        <% End If %>
        <% If (p1.Fields.Item("yin_7").Value) <> ("Y") Then %>
        <tr>
          <td height="25"><table width="100%" border="0" cellspacing="0" cellpadding="0" height="25">
              <tr>
                <td width="22%">&nbsp;</td>
                <td><a href="page_logging.asp?<%= MR_keepNone & MR_joinChar(MR_keepNone) & "id=" & p1.Fields.Item("id").Value %>" class="<%=(p1.Fields.Item("font1").Value)%>"><%=(p1.Fields.Item("title7").Value)%></a></td>
              </tr>
            </table></td>
        </tr>
        <% End If %>
        <tr>
          <td height="40">&nbsp;</td>
        </tr>
        <tr>
          <td height="13"><div align="center"><font color="#000000">访问人次 <span class="pt9"><%=(p1.Fields.Item("hit").Value)%></span></font></div></td>
        </tr>
    </table></td>
    <td width="10" valign="top" height="356" bgcolor="#FFFFFF"><div align="center"><img src="images/2-2.gif" width="1" height="450"></div></td>
    <td width="615" valign="top" height="356" bgcolor="#FFFFFF"><table width="99%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td><img src="images/arrow.gif" width="13" height="13"> <%=(p1.Fields.Item("title_news").Value)%></td>
        </tr>
        <tr>
          <td height="6"></td>
        </tr>
        <tr>
          <td height="1" bgcolor="#CCCCCC"></td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
          <td height="22">&nbsp;</td>
        </tr>
        <tr>
          <td><table width="94%" border="0" cellspacing="0" cellpadding="0" align="center">
              <tr>
                <td><%=(p1.Fields.Item("news_main").Value)%></td>
              </tr>
            </table></td>
        </tr>
        <tr>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td height="40">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
</table>
<table width="778" border="0" cellspacing="0" cellpadding="0" align="center" bgcolor="#999999">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td><div align="center"> <font color="#FFFFFF">电话:<%=(p1.Fields.Item("phone").Value)%>
        <% If (p1.Fields.Item("fax").Value) <> ("") Then %>
        &nbsp;传真:<%=(p1.Fields.Item("fax").Value)%>
        <% End If %>
        <% If (p1.Fields.Item("mobile").Value) <> ("") Then %>
        &nbsp;手机:<%=(p1.Fields.Item("mobile").Value)%>
        <% End If %>
        &nbsp;E-mail:<%=(p1.Fields.Item("email").Value)%> &nbsp;&copy;&nbsp;<%=(p1.Fields.Item("title").Value)%> &nbsp;版权所有</font></div></td>
  </tr>
  <tr>
    <td height="8"><% If (p1.Fields.Item("floatlogo").Value) <> ("00.bmp") Then %>
      <div id="chip" style="position: absolute; visibility: visible">
        <table border="0" cellPadding="4" cellSpacing="0" class="bd" width="60">
          <tbody>
            <tr>
              <td align="middle" class="bg2"><a class="prs1" href="<%=(p1.Fields.Item("floatline").Value)%>" target="_blank"><img border="0" class="bd" src="../photos/<%=(p1.Fields.Item("floatlogo").Value)%>"></a></td>
            </tr>
          </tbody>
        </table>
      </div>
      <script event="onload" for="window" language="JavaScript">
pagestart();
</script>
      <% End If %>
    </td>
  </tr>
</table>
</body>
</html>
<%
p1.Close()
%>
