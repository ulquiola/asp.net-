<%
function remove(spid)					'创建自定义函数
	if isarray(session("arr")) then		'判断session("arr")变量是否为数组
		arr=Session("arr")				'为arr变量赋值
		UB=Ubound(arr,1)				'获取数组的最大可用下标
		if UB=0 then
			session("arr")=""			'清空session变量
		else	
			flag=false     				'标记是否已经删除对应的商品信息
			redim arr_temp(UB-1,2)      '定义临时存储购物车信息的数组
			for i=0 to UB
				if arr(i,0)<>spid then			
					for j=0 to 2
						if flag then     				'当删除一条商品信息后
							arr_temp(i-1,j)=arr(i,j)
						else
							arr_temp(i,j)=arr(i,j)		'为arr_temp(i,j)变量赋值
						end if
					next
				else
					flag=true							'为flag变量赋值
				end if
			next
		end if
	end if
	session("arr")=arr_temp								'为session("arr")变量赋值
end function
'删除指定的记录
if request.QueryString("ID")<>"" then			'判断id值是否为空
	call remove(request.QueryString("ID"))		'调用自定义函数
end if
response.Redirect("cart_see.asp")				'跳转到指定的ASP动态页面
%>