<%@Language=VBScript %>


 
<%
   name=Session("LoginID")
   Date2=Request.Form("year1")
   Date3=Request.Form("month1")
   Riqi1=Date2 & "-" & Date3 & "-" & "1"
   begindate=cdate(riqi1)
   km=Request.Form ("km")
   zkm=Request.Form ("zkm")
   ysmoney1=trim(Request.Form ("money"))
   meno=Request.Form ("meno")
   dep=session("dep")
   cwmoney=0
  Set objConn = Server.CreateObject("ADODB.Connection")
  objConn.Open Application("OledbStr") 
  Set objRst=server.CreateObject ("ADODB.Recordset")
  objRst.LockType=3
  objRst.CursorType=3
  set objRst.activeConnection=objConn
  
        
 objRst.Source="insert into shenzhencwys (bztime,name,kem,fkem,meno,ysmoney,dep,cwmoney) Values ( '" & begindate & "','" & name &"','"& km & "','"& zkm &"', '" & meno & "',convert(money,"& ysmoney1 &"),'"& dep &"',convert(money,"& cwmoney &"))"
 objRst.Open

 Response.Redirect ("ccs_input_index.asp?km="& km) 
%>