   
   addddddd
    <%
   response.write(Application("OledbStr"))
   response.write(Application("UseObject"))
   
   connstr="provider=sqloledb;server=10.101.166.35;database=mastersystem;uid=yanglei;pwd=yanglei123;"
   Set objConn = Server.CreateObject("ADODB.Connection")
   objConn.Open connstr   
   Set obj=server.CreateObject ("ADODB.Recordset")
   obj.LockType=3
   obj.CursorType=3
   set obj.activeConnection=objConn
   sql="Select * From people_infomation"
   obj.Source=sql
   obj.Open
    obj.MoveFirst 
    while not obj.EOF
    %>
<%=obj("personname")%>">,<%=obj("personid")%><br>
    <%
    obj.movenext
    wend
    obj.close
    %>	