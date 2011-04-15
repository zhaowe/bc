<%@Language=VBScript %>


 
<%
  
   km=Request.Form("km")
   fkm=Request.Form("zkm")
   sx=Request.Form ("sx")
   dep=session("dep")
 
  Set objConn = Server.CreateObject("ADODB.Connection")
  objConn.Open Application("OledbStr") 
  Set objRst=server.CreateObject ("ADODB.Recordset")
  objRst.LockType=3
  objRst.CursorType=3
  set objRst.activeConnection=objConn
  
        
 objRst.Source="insert into shenzhencwys_dep (dep,kem,fkem,sx) Values ( '" & dep & "','" & km &"','"& fkm & "','"& sx &"')"
 objRst.Open

 Response.Redirect ("ccs_xtwf.asp") 
%>
