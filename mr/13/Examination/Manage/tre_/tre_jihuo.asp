<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include   file="../../Conn/conn.asp"-->   
<%
  id=Replace(Trim(Request("id")),"'","")
  set rsAdd=server.CreateObject("ADODB.RecordSet")
  rsAdd.open "select * from tb_jiaoyanma",conn,3,2   
  ychar="0,1,2,3,4,5,6,7,8,9,A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z"   '将数字和大写字母组成一个字符串   
  yc=split(ychar,",")   '将字符串生成数组
  for i =1 to 1000
  Randomize
  str=yc(Int((35 * Rnd)))&yc(Int((35 * Rnd)))&yc(Int((35 * Rnd)))&yc(Int((35 * Rnd)))&yc(Int((35 * Rnd)))&yc(Int((35 * Rnd))) &yc(Int((35 * Rnd)))&yc(Int((35 * Rnd)))&yc(Int((35 * Rnd)))&yc(Int((35 * Rnd)))&yc(Int((35 * Rnd)))&yc(Int((35 * Rnd)))'数组一般从0开始读取，所以这里为35*Rnd
  set rs=server.createobject("adodb.recordset")
  sql="select * from tb_jiaoyanma where jiaoyanma='"&str&"'"
  rs.open sql,conn,1,1
	  if rs.eof then
	  rsAdd.addNew   '写入数据库
	  rsAdd("jiaoyanma")=str
	  rsAdd.update
	  end if
  next
  rsadd.close   
  set rsadd = nothing
  rs.close   
  set rs = nothing
  conn.close
  response.Write("<script>alert('成功生成校验码！');history.back();</script>")
%>