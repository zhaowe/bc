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
  
  
  Set objConn = Server.CreateObject("ADODB.Connection")
  objConn.Open Application("OledbStr") 
  Set objRst=server.CreateObject ("ADODB.Recordset")
  objRst.LockType=3
  objRst.CursorType=3
  set objRst.activeConnection=objConn
  
        
 objRst.Source="update shenzhencwys set bztime='" & begindate & "',name='" & name &"',kem='"& km & "',fkem='"& zkm &"',meno='" & meno & "',ysmoney=convert(money,"& ysmoney1 &")  where q='"& session("q") &"'"
 objRst.Open

 Response.Redirect ("ccs_ser.asp?km="& km) %>
