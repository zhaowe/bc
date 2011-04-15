<%@ Language=VBScript %>
<% Set objConn = Server.CreateObject("ADODB.Connection")
  objConn.Open Application("OledbStr") 
  Set objRst=server.CreateObject ("ADODB.Recordset")
  objRst.LockType=3
  objRst.CursorType=3
  set objRst.activeConnection=objConn%>
<%year1=trim(Request.Form ("year1"))
  kem=trim(Request.Form ("kem"))
  date1=cdate(year1 & "-" & 1 & "-" & "1")
  date2=cdate(year1 & "-" & 12 & "-" & "1")
 
objrst.Source ="delete from shenzhencwys_je where yea>='"& date1 &"' and yea<='"& date2 &"' and kem='"& kem &"' and dep='"& session("dep") &"'"
objrst.Open 

Response.Redirect ("ccs_yskh_index.asp")


%>

