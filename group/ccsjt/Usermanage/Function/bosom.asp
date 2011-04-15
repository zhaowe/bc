<%
check= Request.Form("text")
functionid=Request.Form("functionid")
if check="edit"then
Response.Redirect"editfunction.asp?"&"which="&functionid
else 
Response.Redirect"delfunction.asp?"&"which="&functionid
end if
%> 