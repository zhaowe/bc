<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
<%
   Set objConn = Server.CreateObject("ADODB.Connection")
   objConn.Open Application("OledbStr") 
   
   
   Set objRst=server.CreateObject ("ADODB.Recordset")
   objRst.LockType=3
   objRst.CursorType=3
   set objRst.activeConnection=objConn
   
   no=trim(Request.Form("TxtNo"))
   objrst.Source ="select userid from userinfo a inner join szairlineuser b " & _
     " on a.loginid=b.logid where b.no='" & no & "'"

   'Response.Write objrst.Source
   'Response.End
   
   objrst.Open 
   if objrst.EOF and objrst.BOF then
      Response.Write "不存在该员工号的用户！"
   else
      dim curuserid
      curuserid=trim(objrst("userid"))
      objrst.Close
      set objrst=nothing
      Response.Redirect "edituser.asp?userid=" & curuserid
      'Response.Write no & "___" & curuserid
   end if 
   
   
%>

</BODY>
</HTML>
