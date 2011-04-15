<%
check= Request.Form("text")
locale=Request.Form("locale")
if check="edit"then
Response.Redirect"editglocale.asp?"&"locale="&locale
else 
Response.Redirect"delgsontable.asp?"&"locale="&locale
end if
%>