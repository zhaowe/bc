<%
If Session("LoginOk")=false then
	Response.Redirect "login.asp"
End if
%>
<!--#include file="DbClass.asp"-->
<%
Dim paraErrNo
paraErrNo = Request.QueryString("ErrNo")
Response.Write "������ʾ��<strong>" & paraErrNo & "</strong>�Ŵ���"
%>
&nbsp;&nbsp;&nbsp;
<a href="UnDelError.asp?ErrNo=<%=paraErrNo%>">�ָ�</a>
&nbsp;&nbsp;&nbsp;
<a href="Restore.asp">������ҳ</a>
<%
Dim strSql
strSql = "Select * from Error Where ErrorNo='" & paraErrNo & "'"
Dim objDML
Set objDML = Server.CreateObject("Com_DML.ClsDML")
Dim iErrNo
on error resume next

Dim objRs
Set objRs = Server.CreateObject("Adodb.Recordset")
set objRs = objDML.ExeSelect(strSql,DbClass)
if Err.number<>0 then
	iErrNo=Err.number
	Err.Clear
	set objDML=nothing
	Response.Redirect "Sorry.asp?ErrNo=" & iErrNo
End If

If Not objRs.EOF Then
%>
<Form id=frmEdit Method="Post">
<input Type="Hidden" name="hidOperate" value="Edit">
<input type="hidden" name="hidErrNo" value="<%=paraErrNo%>">
<p align="center">Error��</>
<div align="center">
  <center>
  <table border="1" width="90%" height="176" cellspacing="1" cellpadding="10">
    <tr>
      <td width="20%" height="18" bgcolor="#0FC4FF" >�ڲ�����ԭ��</td>
      <td width="80%" height="18"><Input Type="Text" name="txtReasonIn" value="<%=objRs("ErrorReasonIn")%>" style="WIDTH: 520px"></td>
    </tr>
    <tr>
      <td width="20%" height="18" bgcolor="#0FC4FF">�ڲ��������</td>
      <td width="80%" height="18"><Input Type="Text" name="txtSolution" value="<%=objRs("ErrorSolutionIn")%>" style="WIDTH: 520px"></td>
    </tr>
    <tr>
      <td width="20%" height="18" bgcolor="#0FC4FF">A���������</td>
      <td width="80%" height="18">
      <Select name="selClassA">
		<option value="1" <% If LCase(objRs("ClassAType")) = "java" Then %> selected <% End If%>>Java</option>
		<option value="2" <% If LCase(objRs("ClassAType")) = "sql" Then %> selected <% End If%>>Sql</option>
		<option value="3" <% If LCase(objRs("ClassAType")) = "vb" Then %> selected <% End If%>>VB</option>
		<option value="4" <% If LCase(objRs("ClassAType")) = "vc" Then %> selected <% End If%>>VC</option>
		<option value="5" <% If LCase(objRs("ClassAType")) = "vi" Then %> selected <% End If%>>Vi</option>
		<option value="6" <% If LCase(objRs("ClassAType")) = "other" Then %> selected <% End If%>>����</option>
	  </Select></td>
    </tr>
    <tr>
      <td width="20%" height="18" bgcolor="#0FC4FF">B���������</td>
      <td width="80%" height="18">
      <%
Dim ClassBobjRs
Dim ClassBobjDML
Dim ClassBstrSql
Dim ClassBiReturnErr

set ClassBobjRs = server.CreateObject("adodb.recordset")
set ClassBobjDML = server.CreateObject("Com_DML.clsDML")

ClassBstrSql = "select * from ErrorTypeB"
set ClassBobjRs = ClassBobjDML.ExeSelect(ClassBstrSql,DbClass)

if Err.number<>0 then
	ClassBiReturnErr = 10051  '��ѯErrorTypeB�����
	Err.Clear
	set objDML=nothing
	Response.Redirect "Sorry.asp?ErrNo=" & ClassBiReturnErr
End If
%>
      <Select name="selClassB">
      <%do while not ClassBobjRs.EOF%>
      <%if objRs("ClassBType") = ClassBobjRs("ClassBType") then%>
		<option value="<%=ClassBobjRs("ClassBType")%>" selected><%=ClassBobjRs("ClassBName")%></option>
	  <%else%>
	    <option value="<%=ClassBobjRs("ClassBType")%>"><%=ClassBobjRs("ClassBName")%></option>	
	  <%
	  end if
	  ClassBobjRs.MoveNext
	  loop
	  ClassBobjRs.close
	  set ClassBobjRs = nothing
	  set ClassBobjDML = nothing
	  %>
	  </Select></td>
    </tr>
    <tr>
      <td width="20%" height="18" bgcolor="#0FC4FF">���������</td>
      <td width="80%" height="18"><Input Type="Text" name="txtErrPrg" value="<%=objRs("ErrorPrgName")%>" style="WIDTH: 520px"></td>
    </tr>
    <tr>
      <td width="20%" height="18" bgcolor="#0FC4FF">���ʱ��ָ��</td>
      <td width="80%" height="18"><Input Type="Text" name="txtErrGoto" value="<%=objRs("ErrorGoto")%>"></td>
    </tr>
    <tr>
      <td width="20%" height="18" bgcolor="#0FC4FF">�ڲ��ⲿ</td>
      <td width="80%" height="18">
      <Select name="selType">
		<option value="1" <% If objRs("ErrorType") = "i" Then %> selected <% End If%>>�ڲ�</option>
		<option value="2" <% If objRs("ErrorType") = "o" Then %> selected <% End If%>>�ⲿ</option>
	  </Select></td>
    </tr>
    
  </table>
  </center>
</div>
<%
'�ر�Error���¼��
	objRs.Close
	Set objRs = Nothing
	Set objDML = Nothing
End If

'************************************
'���ӱ��¼��
Set objDML = Server.CreateObject("Com_DML.ClsDML")
Dim iLocale
Dim strLocale
strLocale = "Select * from LocaleType Where ErrorNo='" & paraErrNo & "'"

Set objRs = Server.CreateObject("Adodb.Recordset")
set objRs = objDML.ExeSelect(strLocale,DbClass)

if Err.number<>0 then
	iLocale=Err.number
	Err.Clear
	set objDML=nothing
	Response.Redirect "Sorry.asp?ErrNo=" & iLocale
End If

do while Not objRs.EOF
	if objRs("LocaleType") = "zh" and objRs("ProtocolType") = "http" then
%>
<p align="center">LocaleType��</>
<div align="center">
  <center>
  <table border="1" width="90%" height="120" cellspacing="1" cellpadding="10">
    <tr>
	  <td width="20%" height="18" bgcolor="#0FC4FF">���԰汾</td>
      <td width="80%" height="18">
      <p>Zh</p>
    </tr>
    <tr>
      <td width="20%" height="18" bgcolor="#0FC4FF">Э������</td>
      <td width="80%" height="18">
      <p >Http</p></td>
    </tr>
    <tr>
      <td width="20%" height="18" bgcolor="#0FC4FF">�ⲿ������ʾ</td>
      <td width="80%" height="18"><Input Type="Text" name="txtNameOut" value="<%=objRs("ErrorNameOut")%>" style="WIDTH: 520px"></td>
    </tr>
    <tr>
      <td width="20%" height="18" bgcolor="#0FC4FF">�ⲿ�������</td>
      <td width="80%" height="18"><Input Type="Text" name="txtSolutionOut" value="<%=objRs("ErrorSolutionOut")%>" style="WIDTH: 520px"></td>
    </tr>
</table>
  </center>
</div>
	<%end if
	objRs.movenext
	loop%>
</form>

<div align="center"><center>
	<table border="0" width="95%">
	<tr>
		<td align="center" width="10%" bgcolor="#7598FF">�����</td>
		<td align="center" width="8%" bgcolor="#7598FF">���԰汾</td>
		<td align="center" width="8%" bgcolor="#7598FF">Э������</td>
		<td align="center" width="30%" bgcolor="#7598FF">�ⲿ������ʾ</td>
		<td align="center" width="30%" bgcolor="#7598FF">�ⲿ����취</td>	
		<td align="center" width="14%" bgcolor="#7598FF">����޸�<br>ʱ��</td>
	</tr>
<%
objRs.movefirst
do while not objRs.EOF
%>
	<tr>
		<td width="10%"><%=objRs("ErrorNo")%></td>
		<td width="8%"><%=objRs("LocaleType")%></td>
		<td width="8%"><%=objRs("ProtocolType")%></td>
		<td width="30%"><%=objRs("ErrorNameOut")%></td>
		<td width="30%"><%=objRs("ErrorSolutionOut")%></td>
		<td width="14%"><%=objRs("LastModify")%></td>
	</tr>
<%
	objRs.movenext
loop
objRs.Close
Set objRs = Nothing
Set objDML = Nothing
%>
</table></center></div>