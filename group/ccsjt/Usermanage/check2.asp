<%@ Language=VBScript %>

<% 


   dim loginid
   dim password
   dim objRs
   Dim ObjUser 
   dim LoginPass
   Dim iErrNo
   dim FuncStr,StrFunc


   session("LoginID")=""
   LoginID=trim(request.form("loginid"))
   Password=trim(request.form("password"))


   Set objConn = Server.CreateObject("ADODB.Connection")
   objConn.Open "provider=sqloledb;server=10.254.0.46;database=mastersystem;uid=sa;pwd=;"
   
   Set objRst=server.CreateObject ("ADODB.Recordset")
   objRst.LockType=3
   objRst.CursorType=3
   set objRst.activeConnection=objConn

   objRst.Source="select * from View_userInfo where LoginID='" & LoginID & "' and Password='" & Password & "'"
   objRst.Open  

   if not (objRst.BOF and objrst.EOF ) then
      loginpass=true
      Response.Write (loginpass)

      Session("UID")=ObjRst("UserID")
      Session("LoginID")=LoginID
      Session("AgentID")=ObjRst("AgentID")
      Session("IntraLoginOk")=true

      Response.Write (Session("UID"))

      
      objRst.close  
      objRst.Source="select OrderNum from View_FuncStr where UserID='" & Session("UID") & "'" & " Order by OrderNum Asc"
      objRst.open
       
      StrFunc = ","
      Do While Not objRst.EOF
        StrFunc = StrFunc & CStr(objRst("OrderNum")) & ","
        objRst.MoveNext
      Loop
      '取用户所属组，并加上该组拥有的功能OrderNum，相同的功能要过滤
      
      objRst.close  
      objRst.Source="Select * from View_GroupFuncStr where UserID='" & Session("UID") & "'"
      objRst.open
      
      Do While Not objRst.EOF
        str1 = CStr(objRst("OrderNum"))
        Str2 = "," & str1 & ","
        CompResult = InStr(1, StrFunc, Str2)
        If CompResult = 0 Then
            StrFunc = StrFunc & str1 & ","
        End If
        objRst.MoveNext
      Loop
      FuncStr = StrFunc

      Session("FuncStr")=FuncStr

      Response.Write (Session("FuncStr"))
      
      
   
   end if
   
   ObjRst.Close 
   set ObjRst=nothing
   Response.Redirect "../public/public.asp"
   
   
   
%>
