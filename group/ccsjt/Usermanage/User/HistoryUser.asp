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
<p><font color="#3333FF"><b><font color="#000000">��ʷ�û���¼��Ϣ</font></b></font> </p>
<table border="0" width="80%" bgcolor="#003333" cellspacing="1" cellpadding="4">
    <tr> 
      <td><font color="#FFFFFF">�û���</font></td>
      <td><font color="#FFFFFF">��ʼʱ��</font></td>
      <td><font color="#FFFFFF">��ֹʱ��</font></td>
      <td><font color="#FFFFFF">�û�״̬</font></td>
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
  
<p>[ <a href="userinfo.asp"><b>����</b></a> ]</p>
  </BODY>
</HTML>
