<HTML>
<HEAD>
<%
On Error Resume Next
dim userid
userid=Request.QueryString("userid")
dim obj
set obj=server.CreateObject ("Com_UserManage1.ClsUserManage1")
dim objrs 
set objrs=server.CreateObject ("Adodb.recordset")
set objrs=obj.GetUserHistory (UserID)
if Err.number <>0 then
ierror=Err.number 
Err.Clear
set obj=nothing
Response.Redirect"../../Sorry.asp?Errorno="&ierror
end if
set obj=nothing

%>
<TITLE></TITLE>
<link rel="stylesheet" href="../../style.css"> 
</HEAD>
<BODY>
<p><font color="#3333FF"><b><font color="#000000">历史用户记录信息</font></b></font> </p>
<table border="0" width="80%" bgcolor="#003333" cellspacing="1" cellpadding="4">
    <tr> 
      <td><font color="#FFFFFF">用户名</font></td>
      <td><font color="#FFFFFF">起始时间</font></td>
      <td><font color="#FFFFFF">终止时间</font></td>
      <td><font color="#FFFFFF">用户状态</font></td>
    </tr>
    <%
    objrs.MoveFirst 
    do while not objrs.EOF and not objrs.BOF 
    %> 
    <TR bgcolor="#FFFFFF"> 
      <TD height="23"><%=OBJRS("loginid")%></TD>
      <TD height="23"><%=objrs("startdate")%></TD>
      <TD height="23"><%=objrs("enddate")%></TD>
      <TD height="23"><%=objrs("status")%></TD>
    </TR>
    <%
    objrs.MoveNext 
    loop
    %> 
  </table>
  
<p>[ <a href="userinfo.asp"><b>返回</b></a> ]</p>
  </BODY>
</HTML>
