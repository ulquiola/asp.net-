<%
'---------- ����ת���ַ����� -----------
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
'---------- ����ʱ���ȡ���ַ��� -----------
Function GetOrderNo(dDate)
    GetOrderNo = RIGHT("0000"+Trim(Year(dDate)),4)+RIGHT("00"+Trim(Month(dDate)),2)+RIGHT("00"+Trim(Day(dDate)),2)+RIGHT("00" + Trim(Hour(dDate)),2)+RIGHT("00"+Trim(Minute(dDate)),2)+RIGHT("00"+Trim(Second(dDate)),2)
End Function
'---------- ��ȡ��������� -----------
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
//----------- ��֤���� ----------
function checkNum(Num){
  var Expression=/^[1-9]+[0-9]*\d$/;
  var re=new RegExp(Expression);
  if(re.test(Num)==true){
    return true;}
  else{
    return false;}
} 
//----------- ��֤E-mail��ַ ---------- 
function checkEmail(email){
  var Expression=/\w+([-+.']\w+)*\.\w+([-.]\w+)*/;
  var re=new RegExp(Expression);
  if(re.test(email)==true){
    return true;}
  else{
    return false;}
}
//----------- ��֤��ַ ---------- 
function checkUrl(url){
  var Expression=/http(s)?:\/\/([\w-]+\.)+[\w-]+(\/[\w-.\/?%&=]*)?/;
  var re=new RegExp(Expression);
  if(re.test(url)==true){
    return true;}
  else{
    return false;}
}
//----------- ��֤���֤���� ----------
function checkCode(code) {
  var Expression=/\d{17}[\d|X]|\d{15}/;
  var re=new RegExp(Expression);
  if(re.test(code)==true){
    return true;}
  else{
    return false;}
}
</script>
