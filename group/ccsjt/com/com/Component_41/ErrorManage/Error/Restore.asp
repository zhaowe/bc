<%
If Session("LoginOk")=false then
	Response.Redirect "login.asp"
End if
%>
<html>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Sub btnStartQuery_onclick
	frmOperate.action = "restore.asp?strField=" & frmOperate.selField.value & "&strFlag=" & frmOperate.selFlag.value & "&strValue=" & frmOperate.txtValue.value
	frmOperate.submit
End Sub
-->
</SCRIPT>
<%
const PAGE_SIZE = 5
Dim objDML
Set objDML = Server.CreateObject("Com_ErrorManage.clsErrorManage")
Dim iErrNo
on error resume next
Dim objRs
Set objRs = Server.CreateObject("Adodb.Recordset")

'��ȡ��ѯ����
If Request.QueryString("strField") = "" Then
	Set objRs = objDML.ErrorQuery("UserId",3,Session("UserId"),"y")  'ȡ��¼�û���ӵļ�¼
	'Set objRs = objDML.ErrorQuery()
Else
	'strFlagȡֵ:=(1),>(2),like(3),<(4)
	Dim iFlag
	iFlag = cint(Request.QueryString("strFlag"))
	set objRs = objDML.ErrorQuery(Request.QueryString("strField"),iFlag,Request.QueryString("strValue"),"y")
End If
if Err.number<>0 then
	iErrNo = Err.number
	Err.Clear
	set objDML=nothing
	Response.Redirect "Sorry.asp?ErrNo=" & iErrNo
End If

objRs.PageSize = PAGE_SIZE
Dim iCurrentPage

if CInt(Request.QueryString("PageNo"))>=1 and CInt(Request.QueryString("PageNo"))<=objRs.PageCount then
	iCurrentPage = CInt(Request.QueryString("PageNo"))
else
	iCurrentPage =1
end if
%>
<body>
<FORM name="frmOperate" Method="Post" Action="operate.asp">
��ѯ������
<TABLE border=1 cellPadding=1 cellSpacing=1 >
    <TR>
        <TD><SELECT id=selField name=selField>
				<OPTION value=ErrorNo>ErrorNo
			    <OPTION value=ErrorReasonIn>ErrorReasonIn
			    <OPTION value=ErrorSolutionIn>ErrorSolutionIn
			    <OPTION value=UserName>UserName
			    <OPTION value=ClassAType>ClassA
			    <OPTION value=ClassBType>ClassB
			    <OPTION value=ErrorNameOut>ErrorNameOut
			    <OPTION value=ErrorSolutionOut>ErrorSolutionOut
                <OPTION value=���м�¼>���м�¼
			</SELECT>             
		</TD>
        <TD><SELECT id=selFlag name=selFlag>
                <%'<OPTION value=1>����
                '<OPTION value=2>����%>                
                <OPTION selected value=3>����
                <%'<OPTION value=4>С��%>
            </SELECT>
 
        </TD>
        <TD>
            <INPUT id=txtValue name=txtValue>
        </TD>
        <TD>
            <INPUT id=btnQuery name=btnStartQuery type=button value="��ʼ����"> 
        </TD>
        </TR></TR>
</TABLE>
<HR>
<%
If not objRs.EOF Then
	objRs.AbsolutePage = iCurrentPage
	If iCurrentPage > 1 Then
		Response.Write "<A href='Operate.asp?PageNo=" & (iCurrentPage-1) & "&strField=" & Request.QueryString("strField") & "&strFlag=" & Request.QueryString("strFlag") & "&strValue=" & Request.QueryString("strValue") & "'>��һҳ</a>&nbsp;&nbsp;"
	End If
	If iCurrentPage < objRs.PageCount Then
		Response.Write "<A href='Operate.asp?PageNo=" & (iCurrentPage+1) & "&strField=" & Request.QueryString("strField") & "&strFlag=" & Request.QueryString("strFlag") & "&strValue=" & Request.QueryString("strValue") & "'>��һҳ</a>&nbsp;&nbsp;"
	End If
	
%>

	�� <%=iCurrentPage%> / <%=objRs.PageCount%> ҳ<BR><BR>
	<div align="center"><center>
	<table border="0" width="200%">
		<tr>
		<td align="center" width="5%" bgcolor="#7598FF">�����</td>
		<td align="center" width="4%" bgcolor="#7598FF">�û���</td>
		<td align="center" width="10%" bgcolor="#7598FF">�ڲ�����ԭ��</td>
		<td align="center" width="10%" bgcolor="#7598FF">�ڲ�����취</td>
		<td align="center" width="10%" bgcolor="#7598FF">�ⲿ������ʾ</td>
		<td align="center" width="5%" bgcolor="#7598FF">�����</td>
		<td align="center" width="10%" bgcolor="#7598FF">�ⲿ������</td>
		<td align="center" width="2%" bgcolor="#7598FF">A��</td>
		<td align="center" width="6%" bgcolor="#7598FF">B��</td>
		<td align="center" width="8%" bgcolor="#7598FF">���������</td>
		<td align="center" width="2%" bgcolor="#7598FF">����</td>
		<td align="center" width="8%" bgcolor="#7598FF">����޸�����</td>
		</tr>
<%
	dim i
	For i=1 to PAGE_SIZE
%> 
		<tr>
		<td align="center" width="5%"><a href="DelTag.asp?ErrNo=<%=objRs("ErrorNo")%>"><%=objRs("ErrorNo")%></td>
		<td align="center" width="4%"><%=objRs("UserName")%></td>
		<td align="center" width="10%"><%=objRs("ErrorReasonIn")%></td>
		<td align="center" width="10%"><%=objRs("ErrorSolutionIn")%></td>
		<td align="center" width="10%"><%=objRs("ErrorNameOut")%></td>
		<td align="center" width="5%"><a href="DelTag.asp?ErrNo=<%=objRs("ErrorNo")%>"><%=objRs("ErrorNo")%></td>
		<td align="center" width="10%"><%=objRs("ErrorSolutionOut")%></td>
		<td align="center" width="2%"><%=objRs("ClassAType")%></td>
		<td align="center" width="6%"><%=objRs("ClassBName")%></td>
		<td align="center" width="8%"><%=objRs("ErrorPrgName")%></td>
		<td align="center" width="2%"><%=objRs("ErrorType")%></td>
		<td align="center" width="8%"><%=left(objRs("LastModify"),10)%></td>
		</tr>
<%	
		objRs.movenext
		If objRs.EOF Then
			Exit For
		End If
	next%>
	</table></center></div>
<%
Else
	Response.Write "��ǰû�м�¼"
End If
set objDML=nothing
%>
&nbsp;&nbsp;<A href="Operate.asp">����</A>
</form>
</head></html>