<%
'---------- 定义转换字符函数 -----------
Function Str_filter(InString)
  NewStr=Replace(InString,"'","''")
  NewStr=Replace(NewStr,"<","&lt")
  NewStr=Replace(NewStr,">","&gt")
  NewStr=Replace(NewStr,"chr(60)","&lt;")
  NewStr=Replace(NewStr,"chr(37)","&gt;")
  NewStr=Replace(NewStr,"""","&quot")
  NewStr=Replace(NewStr,";",";;")
  NewStr=Replace(NewStr,"--","-")
  NewStr=Replace(NewStr,"/*"," ")
  NewStr=Replace(NewStr,"%"," ")
  Str_filter=NewStr
End Function
'---------- 根据时间获取的字符串 -----------
Function GetOrderNo(dDate)
    GetOrderNo = RIGHT("0000"+Trim(Year(dDate)),4)+RIGHT("00"+Trim(Month(dDate)),2)+RIGHT("00"+Trim(Day(dDate)),2)+RIGHT("00" + Trim(Hour(dDate)),2)+RIGHT("00"+Trim(Minute(dDate)),2)+RIGHT("00"+Trim(Second(dDate)),2)
End Function
'---------- 获取随机数函数 -----------
Function randStr(num)
  strings="0,1,2,3,4,5,6,7,8,9,A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z"
  str=split(strings,",")
  for i=1 to num
  Randomize
  str1=str(int((ubound(str))*rnd))
  returnstr=returnstr&str1
  next
  randStr=returnstr
End Function
%>
<script language="javascript">
//----------- 验证数字 ----------
function checkNum(Num){
  var Expression=/^[1-9]+[0-9]*\d$/;
  var re=new RegExp(Expression);
  if(re.test(Num)==true){
    return true;}
  else{
    return false;}
} 
//----------- 验证E-mail地址 ---------- 
function checkEmail(email){
  var Expression=/\w+([-+.']\w+)*\.\w+([-.]\w+)*/;
  var re=new RegExp(Expression);
  if(re.test(email)==true){
    return true;}
  else{
    return false;}
}
//----------- 验证网址 ---------- 
function checkUrl(url){
  var Expression=/http(s)?:\/\/([\w-]+\.)+[\w-]+(\/[\w-.\/?%&=]*)?/;
  var re=new RegExp(Expression);
  if(re.test(url)==true){
    return true;}
  else{
    return false;}
}
//----------- 验证身份证号码 ----------
function checkCode(code) {
  var Expression=/\d{17}[\d|X]|\d{15}/;
  var re=new RegExp(Expression);
  if(re.test(code)==true){
    return true;}
  else{
    return false;}
}
</script>
