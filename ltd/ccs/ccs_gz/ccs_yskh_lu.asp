<%@Language=VBScript %>


 
<%
   name=Session("LoginID")
   redim m(12) 
   redim datet(12)  
   riqi=Request.Form("year1")
   kem=Request.Form("kem")
   for i=1 to 12
   m(i)=Request.Form ("m"&i)
   next
   
   for i=1 to 12
   datet(i)=cdate(riqi & "-" & i & "-" & "1")
   Response.Write datet(i) 
   next
   
  Set objConn = Server.CreateObject("ADODB.Connection")
  objConn.Open Application("OledbStr") 
  Set objRst=server.CreateObject ("ADODB.Recordset")
  objRst.LockType=3
  objRst.CursorType=3
  set objRst.activeConnection=objConn
 for i=1 to 12 
 if m(i)<>""  then
 objrst.Source ="select * from shenzhencwys_je where kem='"& kem &"' and yea='"& datet(i) &"'"
 objrst.Open 
 if objrst.EOF and objrst.BOF then
 objrst.Close 
 objRst.Source="insert into shenzhencwys_je (kem,mon,yea,name,dep) Values ( '" & kem & "',convert(money,"& m(i) &") ,'"& datet(i) & "','"& name &"','"&  session("dep") &"')"
 objRst.Open
 else
  objrst.Close 
 objrst.Source ="update shenzhencwys_je set mon=convert(money,"& m(i) &") where  dep='"& session("dep") &"' and kem='"& kem &"' and yea='"& datet(i) &"'" 
 objrst.Open 

 end if
 end if 
 next
 Response.Redirect "ccs_yskh_index.asp" 
 %> 

<br>
<a href="glindex.asp">·µ»Ø</a>
