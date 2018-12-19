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
hit.CoMMandText = "UPDATE website  SET hit =hit+1 WHERE id ='" + Replace(hit__MRColParam, "'", "''") + "'"
hit.CoMMandType = 1
hit.CoMMandTimeout = 0
hit.Prepared = true
hit.Execute()
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
<html>
<head>
<title><%=(p1.Fields.Item("title").Value)%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link 
href="../website/templet2/css/css.css" rel=stylesheet type="text/css">
<style type="text/css">
<!--
.pt9 {  font-family: "宋体"; font-size: 9pt; line-height: 14pt}
-->
</style>
</head>
<% If (p1.Fields.Item("floatlogo").Value) <> ("00.bmp") Then %>
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
  {
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
<% End If %>
<body bgcolor="#EFEFEF" text="#000000" leftmargin="0" topmargin="0">
<table width="776" border="0" cellspacing="0" cellpadding="0" height="60" align="center" background="images/back.gif">
  <tr>
    <td><% If (p1.Fields.Item("logo").Value) <> ("0.bmp") Then 'script %>
      <img src="../photos/<%=(p1.Fields.Item("logo").Value)%>" width="<%=(p1.Fields.Item("logo_w").Value)%>" height="<%=(p1.Fields.Item("logo_h").Value)%>">
      <% End If %></td>
    <td><% If p1.Fields.Item("banner").Value = ("000.bmp") Then %>
      <div align="right"; style="font size=<%=(p1.Fields.Item("titlesize").Value)%>; FILTER: <%=(p1.Fields.Item("cssopen").Value)%>(color=<%=(p1.Fields.Item("titlecss").Value)%>,direction=<%=(p1.Fields.Item("titleway").Value)%>); WIDTH: 100%; COLOR: <%=(p1.Fields.Item("titlecolor").Value)%>; FONT-FAMILY: <%=(p1.Fields.Item("titlefont").Value)%>"><b><%=(p1.Fields.Item("title").Value)%>&nbsp;</b></div>
      <% End If %>
      <% If p1.Fields.Item("banner").Value <> ("000.bmp") Then 'script %>
      <div align="right"><img src="../photos/<%=(p1.Fields.Item("banner").Value)%>" width="<%=(p1.Fields.Item("banner_w").Value)%>" height="<%=(p1.Fields.Item("banner_h").Value)%>"></div>
      <% End If %></td>
  </tr>
</table>
<table width="776" border="0" cellspacing="0" cellpadding="0" align="center" bgcolor="#666666">
  <tr>
    <td bgcolor="#333333" height="1"></td>
  </tr>
  <tr>
    <td bgcolor="#94aad6" height="2"></td>
  </tr>
  <tr>
    <td bgcolor="#94aad6" height="22"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr bgcolor="#FF9900">
          <td width="126"><font color="#000000"></font></td>
          <td width="452"><marquee>
            <font color="#000000">
            <% If (p1.Fields.Item("yin_zimu").Value) <> ("Y") Then%>
            <span class="pt9"><%=(p1.Fields.Item("zimu").Value)%></span>
            <% End If %>
            </font>
            </marquee></td>
          <td width="188"><div align="right">
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
              &nbsp; </div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td bgcolor="#999999" height="1"></td>
  </tr>
</table>
<table width="776" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr valign="top">
    <td width="128" bgcolor="#e6e6e6" background="../website/templet2/images/cpjs.gif"><table width="128" border="0" cellspacing="0" cellpadding="0">
        <% If (p1.Fields.Item("yin_1").Value) <> ("Y") Then %>
        <tr>
          <td width="128" height="36" background="images/a.gif"><table width="128" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="20">&nbsp;</td>
                <td><div align="center" class="<%=(p1.Fields.Item("font1").Value)%>"><A HREF="page1.asp?<%= MR_keepNone & MR_joinChar(MR_keepNone) & "id=" & p1.Fields.Item("id").Value %>"><%=(p1.Fields.Item("title1").Value)%></A></div>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <% End If %>
        <% If (p1.Fields.Item("yin_news").Value) <> ("Y") Then %>
        <tr>
          <td width="128" height="36" background="images/a.gif"><table width="128" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="20">&nbsp;</td>
                <td><div align="center" class="<%=(p1.Fields.Item("font1").Value)%>"><A HREF="page_news.asp?<%= MR_keepNone & MR_joinChar(MR_keepNone) & "id=" & p1.Fields.Item("id").Value %>"><%=(p1.Fields.Item("title_news").Value)%></A></div>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <% End If %>
        <% If (p1.Fields.Item("yin_3").Value) <> ("Y") Then 'script %>
        <tr>
          <td width="128" height="36" background="images/a.gif"><table width="128" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="20">&nbsp;</td>
                <td><div align="center" class="<%=(p1.Fields.Item("font1").Value)%>"><A HREF="page3.asp?<%= MR_keepNone & MR_joinChar(MR_keepNone) & "id=" & p1.Fields.Item("id").Value %>"><%=(p1.Fields.Item("title3").Value)%></A></div>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <% End If %>
        <% If (p1.Fields.Item("yin_4").Value) <> ("Y") Then %>
        <tr>
          <td width="128" height="36" background="images/a.gif"><table width="128" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="20">&nbsp;</td>
                <td><div align="center" class="<%=(p1.Fields.Item("font1").Value)%>"><A HREF="page4.asp?<%= MR_keepNone & MR_joinChar(MR_keepNone) & "id=" & p1.Fields.Item("id").Value %>"><%=(p1.Fields.Item("title4").Value)%></A></div>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <% End If %>
        <% If (p1.Fields.Item("yin_5").Value) <> ("Y") Then %>
        <tr>
          <td width="128" height="36" background="images/a.gif"><table width="128" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="20">&nbsp;</td>
                <td><div align="center" class="<%=(p1.Fields.Item("font1").Value)%>"><A HREF="page5.asp?<%= MR_keepNone & MR_joinChar(MR_keepNone) & "id=" & p1.Fields.Item("id").Value %>"><%=(p1.Fields.Item("title5").Value)%></A></div>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <% End If %>
        <% If (p1.Fields.Item("yin_6").Value) <> ("Y") Then %>
        <tr>
          <td width="128" height="36" background="images/a.gif"><table width="128" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="20">&nbsp;</td>
                <td><div align="center" class="<%=(p1.Fields.Item("font1").Value)%>"><A HREF="page6.asp?<%= MR_keepNone & MR_joinChar(MR_keepNone) & "id=" & p1.Fields.Item("id").Value %>"><%=(p1.Fields.Item("title6").Value)%></A></div>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <% End If %>
        <% If (p1.Fields.Item("yin_7").Value) <> ("Y") Then %>
        <tr>
          <td width="128" height="36" background="images/a.gif"><table width="128" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="20">&nbsp;</td>
                <td><div align="center" class="<%=(p1.Fields.Item("font1").Value)%>"><A HREF="page_logging.asp?<%= MR_keepNone & MR_joinChar(MR_keepNone) & "id=" & p1.Fields.Item("id").Value %>"><%=(p1.Fields.Item("title7").Value)%></A></div>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <% End If %>
        <tr>
          <td width="128" height="36"><table border="0" cellspacing="0" cellpadding="2" align="center" width="100%">
              <tr>
                <td height="40">&nbsp;</td>
              </tr>
              <tr>
                <td><div align="center" class="pt9">访问人次 <font color="#FF0000"><%=(p1.Fields.Item("hit").Value)%></font></div>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td width="128" height="100">&nbsp;</td>
        </tr>
      </table>
    </td>
    <td width="648" bgcolor="#FFFFFF"><table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td height="42"><table width="100%" border="0" cellspacing="0" cellpadding="0" height="42">
              <tr>
                <td width="160">&nbsp;</td>
                <td><table width="254" border="0" cellspacing="0" cellpadding="0" height="42">
                    <tr>
                      <td><table width="100%" border="0" cellspacing="0" cellpadding="0" height="42">
                          <tr>
                            <td width="36%"></td>
                            <td valign="bottom"><div align="left"><b class="pt105"><font color="#3366FF"><%=(p1.Fields.Item("title5").Value)%></font></b></div>
                            </td>
                          </tr>
                          <tr>
                            <td width="36%" height="17"></td>
                            <td height="17"></td>
                          </tr>
                        </table>
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td bgcolor="#FFB240" height="1"></td>
        </tr>
        <tr>
          <td height="22">&nbsp;</td>
        </tr>
        <tr>
          <td><%=(p1.Fields.Item("main5").Value)%></td>
        </tr>
        <tr>
          <td height="40">&nbsp;</td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="776" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td height="22" background="../website/templet2/images/line1.gif" bgcolor="#A6B0DD"><table width="776" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td bgcolor="#CCCCCC" height="17"><div align="right">电话:<%=(p1.Fields.Item("phone").Value)%>
              <% If (p1.Fields.Item("fax").Value) <> ("") Then %>
              &nbsp;传真:<%=(p1.Fields.Item("fax").Value)%>
              <% End If %>
              <% If (p1.Fields.Item("mobile").Value) <> ("") Then %>
              &nbsp;手机:<%=(p1.Fields.Item("mobile").Value)%>
              <% End If %>
              &nbsp;E-mail:<%=(p1.Fields.Item("email").Value)%> &nbsp;&copy;&nbsp;<%=(p1.Fields.Item("title").Value)%> &nbsp;版权所有</div>
          </td>
          <td width="20" bgcolor="#CCCCCC" height="17"></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td height="4" bgcolor="#FFFFFF" background="../website/templet2/images/line1_1.gif"></td>
  </tr>
</table>
<% If (p1.Fields.Item("floatlogo").Value) <> ("00.bmp") Then %>
<div id="chip" style="position: absolute; visibility: visible">
  <table border="0" cellPadding="4" cellSpacing="0" class="bd" width="60">
    <tbody>
      <tr>
        <td align="middle" class="bg2"><a class="prs1" href="<%=(p1.Fields.Item("floatline").Value)%>" target="_blank"><img border="0" class="bd" src="../../website/photos/<%=(p1.Fields.Item("floatlogo").Value)%>"></a></td>
      </tr>
    </tbody>
  </table>
</div>
<script event="onload" for="window" language="JavaScript">
pagestart();
</script>
<% End If %>
</body>
</html>
<%
p1.Close()
%>
