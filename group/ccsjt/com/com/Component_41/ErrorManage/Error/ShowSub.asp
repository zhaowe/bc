<%
If Session("LoginOk")=false then
	Response.Redirect "login.asp"
End if
%>
<!--#include file="DbClass.asp"-->
<%
Dim paraErrNo
Dim paraOperate
paraOperate = Request.QueryString("Operate")
paraErrNo = Request.QueryString("ErrNo")
Response.Write "正在显示<strong>" & paraErrNo & "</strong>号错误的子表<BR><BR>"
Dim objDML
Set objDML = Server.CreateObject("Com_DML.ClsDML")
Dim iErrNo
Dim strSql
on error resume next
strSql = "Select * from LocaleType Where ErrorNo='" & paraErrNo & "'"
Dim objRs
Set objRs = Server.CreateObject("Adodb.Recordset")
Set objRs = objDML.ExeSelect(strSql,DbClass)

if Err.number<>0 then
	iErrNo = Err.number
	Err.Clear
	set objDML=nothing
	Response.Redirect "Sorry.asp?ErrNo=" & iErrNo
End If
%>
	<div align="center"><center>
	<table border="0" width="95%">
		<tr>
		<td width="8%" bgcolor="#7598FF">语言类型</td>
		<td width="8%" bgcolor="#7598FF">访问方式</td>
		<td width="42%" bgcolor="#7598FF">外部错误显示</td>
		<td width="42%" bgcolor="#7598FF">外部解决办法</td>
		</tr>
<%
'*********************************
'         显示已有记录
'*********************************
do while not objRs.EOF
%>
		<tr>
		<td width="8%"><%=objRs("LocaleType")%></a></td>
		<td width="8%"><%=objRs("ProtocolType")%></td>
		<td width="42%"><%=objRs("ErrorNameOut")%></td>
		<td width="42%"><%=objRs("ErrorSolutionOut")%></td>
		</tr>
		<%
		objRs.movenext
		loop
		%>
	</table></center></div>
<%
objRs.Close
set objRs=nothing
%>
<%
'*****************************
'   判断是添加还是修改记录
'*****************************
if paraOperate="edit" then
'**************
'	修改记录
'**************
%>
<%
Dim paraLocale
Dim paraProtocol
'paraErrNo = Request.QueryString("ErrNo")
paraLocale = Request.QueryString("Locale")
paraProtocol = Request.QueryString("Protocol")
strSql = "Select * from LocaleType Where ErrorNo='" & paraErrNo & "' and LocaleType='" & _
         paraLocale & "' and ProtocolType='" & paraProtocol & "'"

Set objRs = Server.CreateObject("Adodb.Recordset")
Set objRs = objDML.ExeSelect(strSql,DbClass)
if Err.number<>0 then
	iErrNo = Err.number
	Err.Clear
	set objDML=nothing
	Response.Redirect "Sorry.asp?ErrNo=" & iErrNo
End If

If Not objRs.EOF Then
%>
<Form id=frmEdit Method="Post" Action="SaveType.asp">
<input Type="Hidden" name="hidOperate" value="Edit">
<input type="hidden" name="hidErrNo" value="<%=paraErrNo%>">
<input type="hidden" name="hidLocale" value="<%=paraLocale%>">
<input type="hidden" name="hidProtocol" value="<%=paraProtocol%>">
<div align="center">
  <center>
  <table border="1" width="90%" height="176" cellspacing="1" cellpadding="10">
    <tr>
      <td width="28%" height="18" bgcolor="#0FC4FF">语言类型</td>
      <td width="72%" height="18">
      <%=paraLocale%>
      </td>
    </tr>
    <tr>
      <td width="28%" height="18" bgcolor="#0FC4FF">访问方式</td>
      <td width="72%" height="18">
      <%=paraProtocol%>
      </td>
    </tr>
    <tr>
      <td width="28%" height="18" bgcolor="#0FC4FF">外部错误显示</td>
      <td width="72%" height="18">
      <Input Type="Text" name="txtNameOut" value="<%=objRs("ErrorNameOut")%>" style="WIDTH: 520px"></td>
    </tr>
    <tr>
      <td width="28%" height="18" bgcolor="#0FC4FF">外部解决办法</td>
      <td width="72%" height="18">
      <Input Type="Text" name="txtSolution" value="<%=objRs("ErrorSolutionOut")%>" style="WIDTH: 520px"></td>
    </tr>
  </table>
  </center>
</div>
<center>
<a href="DelLocale.asp?ErrNo=<%=paraErrNo%>&Locale=<%=objRs("LocaleType")%>&Protocol=<%=objRs("ProtocolType")%>">删除</a>
&nbsp;&nbsp;
<a href="EditError.asp?ErrNo=<%=paraErrNo%>">返回</a>
&nbsp;&nbsp;
<Input Type="Submit" name="btnSubmit" value="提交">
</center>
</form>
<%
end if
objRs.Close
Set objRs = Nothing
Set objDML = Nothing
%>
<%
'**************
'  添加新记录
'**************
else%>
<Form id=frmEdit Method="Post" Action="SaveType.asp">
<input Type="Hidden" name="hidOperate" value="Add">
<input type="hidden" name="hidErrNo" value="<%=paraErrNo%>">
<div align="center">
  <center>
  <table border="1" width="90%" height="176" cellspacing="1" cellpadding="10">
    <tr>
      <td width="28%" height="18" bgcolor="#0FC4FF">语言类型</td>
      <td width="72%" height="18">
      <Select name="selLocale">
		<option value="1">EN</option>
		<option value="2">ZH</option>
		<option value="3">ZH-Hk</option>
	  </Select></td>
    </tr>
    <tr>
      <td width="28%" height="18" bgcolor="#0FC4FF">访问方式</td>
      <td width="72%" height="18">
      <Select name="selProtocol">
		<option value="1">HTTP</option>
		<option value="2">WAP</option>
      </Select></td>
    </tr>
    <tr>
      <td width="28%" height="18" bgcolor="#0FC4FF">外部错误显示</td>
      <td width="72%" height="18">
      <Input Type="Text" name="txtNameOut" value="" style="WIDTH: 520px"></td>
    </tr>
    <tr>
      <td width="28%" height="18" bgcolor="#0FC4FF">外部解决办法</td>
      <td width="72%" height="18">
      <Input Type="Text" name="txtSolution" value="" style="WIDTH: 520px"></td>
    </tr>
  </table>
  </center>
</div>
<center>
<Input Type="Submit" name="btnSubmit" value="提交">
</center>
</form>
<%end if%>