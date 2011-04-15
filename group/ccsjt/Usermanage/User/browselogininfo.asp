<!--#include file="dbclass.asp"-->
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<%on error resume next
dim obj
set obj=server.CreateObject ("Com_UserManage1.clsUserManage1")
dim objrs
set objrs=server.CreateObject ("adodb.recordset")
set objrs=obj.GetVolidLogin(useobject)
if Err.number <>0 then
ierror=Err.number 
Err.Clear 
set obj=nothing
Response.Redirect "../../Sorry.asp?Errorno="&ierror
end if

%> 
<link rel="stylesheet" href="../../style.css">
<body bgcolor="#FFFFFF">
<p><b>指定语言所有无效的登录用户信息记录集</b></p>
<p><%
const PAGE_SIZE = 5
objrs.PageSize = PAGE_SIZE
Dim iCurrentPage

if CInt(Request.QueryString("PageNo"))>=1 and CInt(Request.QueryString("PageNo"))<=objrs.PageCount then
	iCurrentPage = CInt(Request.QueryString("PageNo"))
else
	iCurrentPage =1
end if

If not objrs.EOF Then
	objrs.AbsolutePage = iCurrentPage
	If iCurrentPage > 1 Then
		Response.Write  "<A href='Browselogininfo.asp?PageNo=" & (iCurrentPage-1)  &  "'>上一页</a>&nbsp;&nbsp;"
	End If
	If iCurrentPage < objrs.PageCount Then
		Response.Write "<A href='Browselogininfo.asp?PageNo=" & (iCurrentPage+1) &"'>下一页</a>&nbsp;&nbsp;"
	End If
%> 第 <%=iCurrentPage%> / <%=objrs.PageCount%> 页<BR>
</p>
<table width="610" border="0" bgcolor="#000000" cellspacing="1" cellpadding="4">
    <tr bgcolor="#003333"> 
      <td width="7%" height="24"><font color="#FFFFFF">用户名</font></td>
      <td width="5%" height="24"><font color="#FFFFFF">性别</font></td>
      <td width="8%" height="24"><font color="#FFFFFF">代理商</font></td>
      <td width="10%" height="24"><font color="#FFFFFF">公司名称</font></td>
      <td width="10%" height="24"><font color="#FFFFFF">联系方式</font></td>
      <td width="9%" height="24"><font color="#FFFFFF">开始时间</font></td>
      <td width="10%" height="24"><font color="#FFFFFF">结束时间</font></td>
      <td width="9%" height="24"><font color="#FFFFFF">版本</font></td>
    </tr>
    <%
	dim i
	For i=1 to PAGE_SIZE
%> 
    <tr bgcolor="#FFFFFF"> 
      <td width="7%"><%=objrs("loginid")%></td>
      <td width="5%"> <%if objrs("sex")="M" then%> 男 <%else %> 女 <%end if%> </td>
      <td width="8%"><%=objrs("agentname")%></td>
      <td width="10%"><%=objrs("companyname")%></td>
      <td width="10%"><%=objrs("contactinfo")%></td>
      <td width="9%"><%=objrs("startdate")%></td>
      <td width="10%"><%=objrs("enddate")%></td>
      <td width="9%"><%=objrs("locale")%></td>
    </tr>
    <%	
		objrs.movenext
		If objrs.EOF Then
			Exit For
		End If
	next%> 
  </table>
<%
Else
	Response.Write "当前没有记录"
End If
set obj=nothing
objrs.Close
set objrs=nothing
%>
<br>
 [ <b><a href="userinfo.asp">返回</a></b> ] 
</body>