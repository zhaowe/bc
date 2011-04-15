<%
dim functionid
dim locale
dim objdml
dim rs
dim ierrno
   functionid=request("functionid")
   functionname=request("functionname")
   locale=request("locale")
   'Response.Write functionid
   'Response.Write functionname
   'Response.Write locale
   'Response.end
 
set objdml=server.CreateObject("com_usermanage1.clsFunction1")
on error resume next  
 ierror=objdml.AddFunctionLocale(functionid,functionname,locale)
if Err.number<>0 then
	iErrNo = Err.number 
	Err.Clear
	Response.Redirect "../../Sorry.asp?ErrorNo=" & iErrNo
set objdml=nothing	
End If
set objdml=nothing
Response.Redirect "fsontablerun.asp"& "?functionid="& functionid
Response.end
%>
