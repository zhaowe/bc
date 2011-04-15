<%
If Session("LoginOk")=false then
	Response.Redirect "login.asp"
End if
%>
<!--#include file="DbClass.asp"-->
<form id=frmAdd Method="Post" Action="SaveErr.asp">
<Input Type="hidden" name="hidOperate" value="Add">
<p align="center">Error表</>
<div align="center">
  <center>
  <table border="1" width="90%" height="176" cellspacing="1" cellpadding="10">
    <tr>
      <td width="20%" height="18" bgcolor="#0FC4FF">内部错误原因</td>
      <td width="80%" height="18"><Input Type="Text" name="txtReasonIn" value="" style="WIDTH: 520px"></td>
    </tr>
    <tr>
      <td width="20%" height="18" bgcolor="#0FC4FF">内部解决方法</td>
      <td width="80%" height="18"><Input Type="Text" name="txtSolution" value="" style="WIDTH: 520px"></td>
    </tr>
    <tr>
      <td width="20%" height="18" bgcolor="#0FC4FF">A类错误类型</td>
      <td width="80%" height="18">
      <Select name="selClassA">
        <option value="1">Java</option>
		<option value="2">Sql</option>
		<option value="3">VB</option>
		<option value="4">VC</option>
		<option value="5">Vi</option>
		<option value="6">其它</option>
	  </Select></td>
    </tr>
    <tr>
      <td width="20%" height="18" bgcolor="#0FC4FF">B类错误类型</td>
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
	iReturnErr = 10051 '查询ErrorTypeB表出错
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
      <td width="20%" height="18" bgcolor="#0FC4FF">出错程序名</td>
      <td width="80%" height="18"><Input Type="Text" name="txtErrPrg" value="" style="WIDTH: 520px"></td>
    </tr>
    <tr>
      <td width="20%" height="18" bgcolor="#0FC4FF">解决时的指向</td>
      <td width="80%" height="18"><Input Type="Text" name="txtErrGoto" value=""></td>
    </tr>
    <tr>
      <td width="20%" height="18" bgcolor="#0FC4FF">内部外部</td>
      <td width="80%" height="18">
      <Select name="selType">
		<option value="1">内部</option>
		<option value="2">外部</option>
	  </Select></td>
    </tr>
    
  </table>
  </center>
</div>
<p align="center">LocaleType表</>
<div align="center">
  <center>
  <table border="1" width="90%" height="120" cellspacing="1" cellpadding="10">
    <tr>
	  <td width="20%" height="18" bgcolor="#0FC4FF">语言版本</td>
      <td width="80%" height="18">
      <p>Zh(简体中文)</p>
    </tr>
    <tr>
      <td width="20%" height="18" bgcolor="#0FC4FF">协议类型</td>
      <td width="80%" height="18">
      Http</td>
    </tr>
    <tr>
      <td width="20%" height="18" bgcolor="#0FC4FF">外部错误显示</td>
      <td width="80%" height="18"><Input Type="Text" name="txtNameOut" value="" style="WIDTH: 520px"></td>
    </tr>
    <tr>
      <td width="20%" height="18" bgcolor="#0FC4FF">外部解决方法</td>
      <td width="80%" height="18"><Input Type="Text" name="txtSolutionOut" value="" style="WIDTH: 520px"></td>
    </tr>
</table>
<input type="submit" name="btnSubmit" value="提交">
  </center>
</div>
</form>

