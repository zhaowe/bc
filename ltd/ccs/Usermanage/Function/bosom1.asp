<%
check= Request.Form("text")
locale=Request.Form("locale")
if check="edit"then
Response.Redirect"editfsontable.asp?"&"locale="&locale
else 
Response.Redirect"delfsontable.asp?"&"locale="&locale
end if
%> 