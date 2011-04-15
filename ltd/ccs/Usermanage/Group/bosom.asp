<%
check= Request.Form("text")
groupid=Request.Form("groupid")
if check="edit"then
Response.Redirect"editgroup.asp?"&"which="&groupid
else 
Response.Redirect"delgroup.asp?"&"which="&groupid
end if
%>     