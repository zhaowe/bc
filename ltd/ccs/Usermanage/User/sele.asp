<%@ Language=VBScript %>
<%
dim user,text,Flag,Del
user=Request.Form("user")
text=request("text1")
Del=text
if text="History" then
Response.Redirect "historyuser.asp?userid="&user
end if
if text="Edit" then
Response.Redirect "edituser.asp?userid="&user
end if 
if text="Del" then
Response.Redirect "publicinfo.asp?userid=" & user & "&Flag="&Del
end if
%> 