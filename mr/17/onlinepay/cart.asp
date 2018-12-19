<%
function cart(buyID,price,buynumber)			'创建自定义函数
	if not isarray(session("arr")) then			'判断session("arr")变量是否是数组
		dim arr(0,2)							'定义数组
		arr(0,0)=buyID							'为数组赋值
		arr(0,1)=price							'为数组赋值
		arr(0,2)=buynumber
	else
		arr=session("arr")						'为arr变量赋值
		UB=ubound(arr,1)+1						
		redim arr(UB,2)							'定义动态数组
		sessionarr=session("arr")				'为sessionarr变量赋值
		'将新添加的商品信息保存
		arr(UB,0)=buyID
		arr(UB,1)=price
		arr(UB,2)=buynumber
		'应用循环将保存在seesion("arr")数组中的商品信息添加到arr中
		For I = 0 To ubound(sessionarr,1)
		   For J = 0 To 2
			  arr(I, J)=sessionarr(I,J)
		   Next
		Next	
	end if
	session("arr")=arr   	
end function
'调用购物车函数并显示购物车中的数据
'***********************************
'功能：从购物车中移去指定的商品
'调用方法：call remove(要删除商品的ID号)
'***********************************
function remove(spid)
	if isarray(session("arr")) then
		arr=Session("arr")
		UB=Ubound(arr,1)
		if UB=0 then
			session("arr")=""
		else	
			flag=false '标记是否已经删除对应的商品信息
			redim arr_temp(UB-1,2)  '定义临时存储购物车信息的数组
			for i=0 to UB
				if arr(i,0)<>spid then			
					for j=0 to 2
						if flag then  '当删除一条商品信息后
							arr_temp(i-1,j)=arr(i,j)
						else
							arr_temp(i,j)=arr(i,j)
						end if
					next
				else
					flag=true
				end if
			next
		end if
	end if
	session("arr")=arr_temp
end function
%>
