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
	Response.Write "�������:" & "10053"
	Response.Write "<br>"
	Response.Write "<br>"
	Response.Write "����ԭ��:" & "sorry.asp������ô򿪼�¼������"
	Response.Write "<br>"
	Response.Write "<br>"
	Response.Write "����취:" & "���Sorry.asp����Com_DML������������"
	Response.End 
end if
Response.Write "�������:" & paraErrNo
Response.Write "<br>"
Response.Write "<br>"
Response.Write "����ԭ��:" & objRs("ErrorReasonIn")
Response.Write "<br>"
Response.Write "<br>"
Response.Write "����취:" & objRs("ErrorSolutionIn")
set objRs=nothing
set ObjDML=nothing
%>
