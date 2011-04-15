<%@ Language=VBScript %>


<%
if trim(session("UID"))="" then

   session("errorNo")="000003"
   Response.Redirect "../../sorry/sorry.asp"

 else
' 
'  Get variables from form in CsnTelephone.asp
'
   Session("Password1") = Request.Form("Password1")
   Session("Password2") = Request.Form("Password2")

   Response.Redirect "ChangePassword_Result.asp"

end if 

%>


