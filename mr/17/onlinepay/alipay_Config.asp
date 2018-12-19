<%
      show_url          = "www.XXXX.com"                   '网站的网址
	  seller_email		= ""				'请设置成您自己的支付宝帐户
	  partner			= ""					'支付宝的账户的合作者身份ID
	  key			    = ""	'支付宝的安全校验码

      notify_url			= ""	'付完款后服务器通知的页面 要用 http://格式的完整路径 比如 :http://localhost/192.168.0.8/alipay/Alipay_Notify.aspx
	  return_url			= ""	'付完款后跳转的页面 要用 http://格式的完整路径 比如 :http://localhost/192.168.0.8/alipay/return_Alipay_Notify.asp

	   logistics_fee		= "0.00"				'物流配送费用
	   logistics_payment	= "BUYER_PAY"			'物流配送费用付款方式：SELLER_PAY(卖家支付)、BUYER_PAY(买家支付)、BUYER_PAY_AFTER_RECEIVE(货到付款)
	   logistics_type		= "EXPRESS"				'物流配送方式：POST(平邮)、EMS(EMS)、EXPRESS(其他快递)
	 
	 
'登陆 www.alipay.com 后, 点商户后台,可以看到支付宝安全校验码和合作id,导航栏的下面
	 
%>