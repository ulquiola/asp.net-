<%
dim xb,mx,my
xb=50 '左边距
my=250 '图形高度
mx=600 '图形宽度
'画矩形

'画边框　画XY轴
function DrawFrame()
	DrawBorder()
	DrawLine xb,30,xb,my+50 'Y轴
	DrawLine xb,my+50,xb+mx,my+50 'X轴
end function 

'画边框
function DrawBorder()
	response.Write("<?xml:namespace prefix=v />")
	response.Write("<v:rect style='left:-5;top:-5;width:"&xb+(mx+100)&"px;height:"&my+100&"px'><v:shadow on='t' type='single' color='silver' offset='5px,5px' /></v:rect>")'边框
end function
'画柱形
function DrawRects(arrX,arrY)
	dim maxY,newArrY,nleft

	'arrY = Array(100,200,300,400,500,600,100,200,300,400,500,600)
	'arrX = Array("一月","二月","三月","四月","五月","六月","一月","二月","三月","四月","五月","六月")
	newArrY=arrY

	maxY=Sort(arrY)
	
	maxY=Formatnumber(maxy,2,-1,-1,0)
	xb=len(maxy)*5 '根据最大值设置左边距的距离
	DrawFrame() '画框架
	DrawScale maxy,arrX '画刻度

	for i=0 to UBound(newArrY)
		nWidth=(mx-50)/(UBound(newArrY)+1)/2
		nHeight=my/maxY*newArrY(i)
		
		nleft=mx-(mx-50)/(UBound(newArrY)+1)*((UBound(newArrY)+1)-i-1)'获取X的刻度
		nleft=nleft+xb-nWidth/2'x的刻度＋边距＋柱形的半宽度
		'nleft=650/7*(i+1)+xb
		
		nTop=my+50-my*newArrY(i)/maxY+15

		DrawRect nleft,nTop,nWidth,nHeight,arrX(i)&"：&#13;"&newArrY(i) ' 画矩形
	next
end function

'画扇形
function DrawArc(item_t,item_v)
	r=my/2+25
	dim maxY,newArrY,nleft

	
	dim item_p(10) '各个项目的比例
	dim item_q(10) '各个项目的百分比
	dim sum '项目总和
	dim nLen '项目个数
	dim d '直径
	dim s
	'定义颜色
	dim color1
	dim color2
	'定义角度
	dim angle1
	dim angle2
	
	nLen=UBound(item_t)
	d=r*2
	color1=Array("#d1ffd1","#ffbbbb","#ffe3bb","#cff4f3","#d9d9e5","#ffc7ab","#ecffb7","#ecffb7")
	color2=Array("#00ff00","#ff0000","#ff9900","#33cccc","#666699","#993300","#99cc00","#99cc00")
	
	for i=0 to nLen
		sum=cint(sum)+cint(item_v(i))
	next
	
	for i=0 to nLen
		item_p(i)=item_v(i)/sum
		item_q(i)=FormatNumber(item_p(i)*100,2,-1,-1,0)+"%"

	next
	
	'背景
	DrawBorder()
	'扇形

	angle1=0
	for i=0 to nLen
		angle2=FormatNumber(360*item_p(i),0,-1,-1,0)

		if i=nLen then
			angle2=360-angle1
		end  if
		s=s&"<v:shape title='"&item_t(i)&"："&item_q(i)&"' style='left:"&xb&"px;top:30px;Z-INDEX:1019;position:absolute;width:"&d&";height:"&d&"' coordsize='"&d&","&d&"' strokeweight='1' strokecolor='#fff' fillcolor='"&color1(i)&"' path='m "&r&","&r&" ae "&r&","&r&","&r&","&r&","&CStr(65536*angle1)&","&CStr(65536*angle2)&" x e'>"
		s=s&"<v:fill color2='"&color2(i)&"' rotate='t' focus='100%' type='gradient' />"
		s=s&"</v:shape>"
		angle1=angle1+angle2
		
	next
	'标注
	s=s&"<v:group style='position:absolute;left:"&(d+25)&";top:"&(d-(22*nLen+12))&";width:200;height:"&(22*nLen+4)&"' coordsize='200,"&(22*nLen+4)&"'>"
	s=s&"<v:rect style='width:200;height:"&(25*nLen+24)&";left:50px' strokecolor='#333' />"

	for i=0 to nLen
		s=s&"<v:rect style='left:4;top:"&(i*22+4)&"px;width:25px;height:15px;;left:55px' title='"&item_t(i)&"："&item_q(i)&"' fillcolor='"&color1(i)&"'><v:fill color2='"&color2(i)&"' rotate='t' focus='100%' type='gradient' /></v:rect>"
		s=s&"<v:shape style='left:30;top:"&(i*22+4)&";width:200;height:25'><v:textbox inset='0,0,0,0' ><table><td style='font-size:10px;padding-left:55px'>"&item_t(i)&"："&item_v(i)&"("&item_q(i)&")</td></table></v:textbox></v:shape>"
	 next
	s=s&"</v:group>"
	
	's=s&"</v:group>"
	response.Write(s)
end function

'画折线
function DrawLines(arrX,arrY)
	dim maxY,newArrY
	'arrY = Array(1,100,10,1,160,20,1,100,10,1,160,20,30)
	'arrX = Array("一月","二月","三月","四月","五月","六月","一月","二月","三月","四月","五月","六月","i其月")
	newArrY=arrY
	maxY=Sort(arrY)
	maxY=Formatnumber(maxy,2,-1,-1,0)
	xb=len(maxy)*5 '根据最大值设置左边距的距离
	DrawFrame() '画框架
	DrawScale maxy,arrX '画刻度
	
	for i=0 to UBound(newArrY)-1
	
		x=GetXScale(UBound(newArrY)+1,(UBound(newArrY)+1)-i-1)+xb
		xn=GetXScale(UBound(newArrY)+1,(UBound(newArrY)+1)-i-2)+xb
		DrawLine x,my+50-my/maxY*newArrY(i),xn,my+50-my/maxY*newArrY(i+1)'画折线
		
		nleft=mx-(mx-50)/(UBound(newArrY)+1)*((UBound(newArrY)+1)-i-1)-3'获取X的刻度
		nleft=nleft+xb-nWidth/2'x的刻度＋边距＋柱形的半宽度
		nTop=my+50-my*newArrY(i)/maxY+15
		DrawOval nleft,nTop,arrX(i)&"：&#13;"&newArrY(i)'描点
		
		if UBound(newArrY)-2=i then '添加最后一个标点
			nleft=mx-(mx-50)/(UBound(newArrY)+1)*((UBound(newArrY)+1)-i-1-2)-3'获取X的刻度
			nleft=nleft+xb-nWidth/2'x的刻度＋边距＋柱形的半宽度
			nTop=my+50-my*newArrY(i+2)/maxY+15
			DrawOval nleft,nTop,arrX(i+2)&"：&#13;"&newArrY(i+2)'描点
		end if
	next
	
end function



'画XY轴的刻度
function DrawScale(maxY,arrX)
	dim yScale,xScale,yLabel
	'画Y轴刻度
	for i=0 to 20
		yScale=my+50-i*(my/20)
		yLabel=GetYStep(maxY)
		yLabel=yLabel*i
		yLabel=Formatnumber(yLabel,2,-1,-1,0) '转换为保留两位小数的格式
		DrawLine xb,yScale,xb+5,yScale
		DrawLabel 30,yScale+9,yLabel
	next
	'画X轴刻度
	for i=1 to UBound(arrX)+1
		'xScale=GetXScale(UBound(arrX),UBound(arrX)-i)+xb
		xScale=GetXScale(UBound(arrX)+1,UBound(arrX)+1-i)+xb
		DrawLine xScale,my+45,xScale,my+50
		DrawLabel xScale+50,my+70,arrX(i-1)
	next
	
end function 

function DrawShape()

end function
'画折线圆点
function DrawOval(nLeft,nTop,title)
	response.Write("<v:oval title='"&title&"' style='Z-INDEX:1034;LEFT:"&nLeft&"px;WIDTH:6px;POSITION:absolute;TOP:"&nTop&"px;HEIGHT:5px' coordsize='21600,21600' fillcolor='red' strokecolor='black' strokeweight='1px'></v:oval>")
end function 
'画直线
function DrawLine(fx,fy,tx,ty)
	response.Write("<v:line style='Z-INDEX:1036;LEFT:50px;POSITION:absolute;' from='"&fx&"px,"&fy&"px' to='"&tx&"px,"&ty&"px' strokecolor='black' strokeweight='1px'></v:line>")
end function

'画柱形
function DrawRect(nLeft,nTop,nWidth,nHeight,title)
	response.write "<v:rect title='"&title&"' style='Z-INDEX:1019;LEFT:"&nLeft&"px;WIDTH:"&nWidth&"px;POSITION:absolute;TOP:"&nTop&"px;HEIGHT:"&nHeight&"px'  fillcolor='#1b72a7' strokecolor='#9d1e1b' strokeweight='1px'><v:shadow on='T' type='single' color='#b3b3b3' offset='5px,5px'/></v:rect>"
end function

'画标注
function DrawLabel(x,y,yScale)
	response.write "<SPAN style='FONT-SIZE:12px;Z-INDEX:1001;LEFT:"&x&"px;COLOR:#000000;FONT-FAMILY:宋体;POSITION:absolute;TOP:"&y&"px;BACKGROUND-COLOR:#ffffff'>"&yScale&"</SPAN>"
end function

'获取X轴刻度
function GetYStep(maxY)
	GetYStep=maxY/20
end function

'获取X轴刻度
function GetXScale(maxX,i)
	GetXScale=(mx-50)-(mx-50)/maxX*i
end function


'冒泡法获取最大值
function Sort(ary) 
dim keepChecking,I,firstValue,secondValue 
keepChecking = true 

do until keepChecking = false 
	keepChecking = false 
	For I = 0 to UBound(ary) 
		if I = UBound(ary) then exit for 
		if cint(ary(I)) < cint(ary(I+1)) then 
				firstValue = ary(I) 
				secondValue = ary(I+1) 
				ary(I) = secondValue 
				ary(I+1) = firstValue 
				keepChecking = true 
		end if 
	next 
Loop 

Sort = ary(0)
end function 
%>