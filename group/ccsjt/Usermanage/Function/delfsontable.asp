
<%
dim functionid
dim functioninfo
dim objdml
dim ierrno
dim localeinfo
dim i
dim howmanyfield
    locale=Request.QueryString ("locale")
    functionid=session("functionid")
    session.Abandon
set objdml=server.createobject("com_usermanage1.clsFunction1")

'on error resume next
    str=objdml.DelFunctionLocale(functionid,locale)
if  err.number<>0 then
    ierrno=err.number
    err=nothing
set objdml=nothing
    response.redirect"../../Sorry.asp?ErrorNo=" & iErrNo
end if 
set objdml=nothing
Response.Redirect"fsontablerun.asp?"&"functionid="&functionid 
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel=stylesheet href="../group/style.css">
<title>ɾ�����ܰ汾��Ϣ</title>
</head>


<body background="../group/images/bg.gif" style="FONT-SIZE: 10.5pt">
<div align=center style="HEIGHT: 43px; WIDTH: 575px"> 
  <p><b><font color="blue" face="����_GB2312" size="6">ɾ�����ܰ汾��Ϣ</font></b> <br>   