<%@Language=VBScript %>


 
<%
  
   km=Request.Form("km")
   dep=session("dep")
 
  Set objConn = Server.CreateObject("ADODB.Connection")
  objConn.Open Application("OledbStr") 
  Set objRst=server.CreateObject ("ADODB.Recordset")
  objRst.LockType=3
  objRst.CursorType=3
  set objRst.activeConnection=objConn
  
        
 objRst.Source="insert into shenzhencwys_khkem (dep,khkem) Values ( '" & dep & "','" & km &"')"
 objRst.Open

 Response.Redirect ("ccs_xtwf_kh.asp") 
%>
