<!--#include file="DbClass.asp"-->
<%
Dim paraErrNo
Dim sqlstr
paraErrNo = Request.QueryString("ErrNo")
sqlstr = "select * from View_ErrorShow where ErrorNo='" & paraErrNo & "'"
set ObjDMl = server.CreateObject("Com_DMl.ClsDML")
set objRs = server.CreateObject("adodb.recordset")
on error resume next
set objRs = ObjDML.ExeSelect(sqlstr,DbClass)
if Err.number <> 0 then 
	Err.Clear
	set objDML=nothing
	Response.Write "错误代码:" & "10053"
	Response.Write "<br>"
	Response.Write "<br>"
	Response.Write "错误原因:" & "sorry.asp本身调用打开记录集出错"
	Response.Write "<br>"
	Response.Write "<br>"
	Response.Write "解决办法:" & "检查Sorry.asp调用Com_DML组件的输入参数"
	Response.End 
end if
Response.Write "错误代码:" & paraErrNo
Response.Write "<br>"
Response.Write "<br>"
Response.Write "错误原因:" & objRs("ErrorReasonIn")
Response.Write "<br>"
Response.Write "<br>"
Response.Write "解决办法:" & objRs("ErrorSolutionIn")
set objRs=nothing
set ObjDML=nothing
%>
