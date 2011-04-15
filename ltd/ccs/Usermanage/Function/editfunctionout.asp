'<!--#include file="public.inc"-->
<%
dim objdml
dim editfunction
dim functionid
dim description
    functionid=session("functionid")
    conflict=Request.Form ("conflict")
    functionname=Request.Form ("functionname")
    ffunctionid=Request.Form("fFunctionid")
    functiontypem=Request.Form("functiontypem")
    functiontypef=Request.Form("functiontypef")
    functiontypep=Request.Form("functiontypep")
    if functiontypem<>""then functiontypem="M" end if
    if functiontypep<>""then functiontypep="P" end if 
    if functiontypef<>""then functiontypeF="F" end if
    functiontype=functiontypem&functiontypef&functiontypep
on error resume next
set objdml=server.CreateObject("Com_UserManage.ClsFunction")
'set editfunction=objdml.EditFunction(functionid,description)
 editfunction=objdml.EditFunctionAll(functionid,ffunctionid,functionname,functiontype,locale)
if Err.number<>0 then 
	iErrNo = Err.number
	Err.Clear
	set objdml=nothing
	Response.Redirect "../../Sorry.asp?ErrorNo=" & iErrNo
end if 
Response.Redirect "functioninfo.asp"
%>
 