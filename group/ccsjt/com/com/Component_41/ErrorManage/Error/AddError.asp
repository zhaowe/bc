<%
If Session("LoginOk")=false then
	Response.Redirect "login.asp"
End if
%>
<!--#include file="DbClass.asp"-->
<form id=frmAdd Method="Post" Action="SaveErr.asp">
<Input Type="hidden" name="hidOperate" value="Add">
<p align="center">Error��</>
<div align="center">
  <center>
  <table border="1" width="90%" height="176" cellspacing="1" cellpadding="10">
    <tr>
      <td width="20%" height="18" bgcolor="#0FC4FF">�ڲ�����ԭ��</td>
      <td width="80%" height="18"><Input Type="Text" name="txtReasonIn" value="" style="WIDTH: 520px"></td>
    </tr>
    <tr>
      <td width="20%" height="18" bgcolor="#0FC4FF">�ڲ��������</td>
      <td width="80%" height="18"><Input Type="Text" name="txtSolution" value="" style="WIDTH: 520px"></td>
    </tr>
    <tr>
      <td width="20%" height="18" bgcolor="#0FC4FF">A���������</td>
      <td width="80%" height="18">
      <Select name="selClassA">
        <option value="1">Java</option>
		<option value="2">Sql</option>
		<option value="3">VB</option>
		<option value="4">VC</option>
		<option value="5">Vi</option>
		<option value="6">����</option>
	  </Select></td>
    </tr>
    <tr>
      <td width="20%" height="18" bgcolor="#0FC4FF">B���������</td>
      <td width="80%" height="18">
<%
Dim objRs
Dim objDML
Dim strSql
Dim iReturnErr

set objRs = server.CreateObject("adodb.recordset")
set objDML = server.CreateObject("Com_DML.clsDML")
on error resume next

strSql = "select * from ErrorTypeB"

set objRs = objDML.ExeSelect(strSql,DbClass)
if Err.number <> 0 then
	iReturnErr = 10051 '��ѯErrorTypeB�����
	Err.Clear
	set objDML=nothing
	Response.Redirect "Sorry.asp?ErrNo=" & iReturnErr
End If
%>
      <Select name="selClassB">
      <%do while not objRs.EOF%>
		<option value="<%=objRs("ClassBType")%>"><%=objRs("ClassBName")%></option>
	  <%
	  objRs.MoveNext
	  loop
	  objRs.close
	  set objRs = nothing
	  set objDML = nothing
	  %>
	  </Select></td>
    </tr>
    <tr>
      <td width="20%" height="18" bgcolor="#0FC4FF">���������</td>
      <td width="80%" height="18"><Input Type="Text" name="txtErrPrg" value="" style="WIDTH: 520px"></td>
    </tr>
    <tr>
      <td width="20%" height="18" bgcolor="#0FC4FF">���ʱ��ָ��</td>
      <td width="80%" height="18"><Input Type="Text" name="txtErrGoto" value=""></td>
    </tr>
    <tr>
      <td width="20%" height="18" bgcolor="#0FC4FF">�ڲ��ⲿ</td>
      <td width="80%" height="18">
      <Select name="selType">
		<option value="1">�ڲ�</option>
		<option value="2">�ⲿ</option>
	  </Select></td>
    </tr>
    
  </table>
  </center>
</div>
<p align="center">LocaleType��</>
<div align="center">
  <center>
  <table border="1" width="90%" height="120" cellspacing="1" cellpadding="10">
    <tr>
	  <td width="20%" height="18" bgcolor="#0FC4FF">���԰汾</td>
      <td width="80%" height="18">
      <p>Zh(��������)</p>
    </tr>
    <tr>
      <td width="20%" height="18" bgcolor="#0FC4FF">Э������</td>
      <td width="80%" height="18">
      Http</td>
    </tr>
    <tr>
      <td width="20%" height="18" bgcolor="#0FC4FF">�ⲿ������ʾ</td>
      <td width="80%" height="18"><Input Type="Text" name="txtNameOut" value="" style="WIDTH: 520px"></td>
    </tr>
    <tr>
      <td width="20%" height="18" bgcolor="#0FC4FF">�ⲿ�������</td>
      <td width="80%" height="18"><Input Type="Text" name="txtSolutionOut" value="" style="WIDTH: 520px"></td>
    </tr>
</table>
<input type="submit" name="btnSubmit" value="�ύ">
  </center>
</div>
</form>

