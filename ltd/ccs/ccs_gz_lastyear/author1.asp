<%

IF trim(session("UID"))<>"" then
  dim objD
   set ObjD=server.CreateObject ("Com_UserManage.ClsUserManage")
       VerifyOk=objD.VerifyUserFunction (session("UID"),"CCS_BMGLY")
   if VerifyOk=false then
     session("errorNo")="000011"
     Response.Redirect "../sorry/sorry.asp"
   end if   
else
  session("errorNo")="000001"
   Response.Redirect "../sorry/sorry.asp"
end if 
  
  

   

%>