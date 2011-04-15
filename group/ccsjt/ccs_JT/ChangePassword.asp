<%@ Language=VBScript %>


<%
   Session("username") = Request.Form("username")
   Session("Password1") = Request.Form("Password1")
   Session("Password2") = Request.Form("Password2")

   Response.Redirect "ChangePassword_Result.asp"


%>


