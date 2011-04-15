<!--#include file="public.inc"-->
<%
dim groupinfo
dim groupid
dim objdml
dim groupuser
dim howmanyfield
   groupid=Request.QueryString("which")
  session("groupid")=groupid
set objdml=server.CreateObject ("Com_UserManage1.clsUserManage1") 
on error resume next
set groupinfo=server.CreateObject("adodb.recordset")
set groupinfo=objdml.GetGroupInfo(groupid)
set objrs=server.createobject("adodb.recordset")
set objrs=objdml.GetAllUser(locale)
if Err.number<>0 then
	iErrNo = Err.number
	Err.Clear
	Response.Redirect "../../Sorry.asp?ErrorNo=" & iErrNo
set objdml=nothing	
End If
set objdml=nothing
howmanyfield=objrs.fields.count-1
dim i
%>    
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>增加组用户管理</a></title>
</head>

<body background="images/bg.gif">
  <div align=center><font size=6 color=blue><strong>增加组用户</strong></font><br>        
  <a href="groupinfo.asp">返回组管理页</a>
  <TABLE WIDTH=75%  >
	<TR>
		<TD width="10%">组名:</TD>
		<TD width="20%"><%=groupinfo(3)%></TD>
		<TD width="10%">描述:</TD>
		<TD><%=groupinfo(1)%></TD>
	</TR>
</TABLE>
</div>
<%
  const MaxPerPage=10
  dim TotalPages
  dim TotalPut
  dim CurrentPage
  if not isempty(request("page"))and isnumeric(request("page")) then
    if request("page")<65025 then
     currentPage=cint(request("page"))
	 else 
     currentPage=1
	 end if
   else
      currentPage=1
   end if
%>


<div align=center>
<%
 
     TotalPut=objrs.recordcount
	 
	 if CurrentPage<1 then
	   CurrentPage=1
	 end if
	 
	 if (CurrentPage-1)*MaxPerPage>TotalPut then
	   if (TotalPut mod MaxPerPage)=0 then
	     CurrentPage=TotalPut \ MaxPerPage
		else
		 CurrentPage=TotalPut \ MaxPerPage + 1
		end if
     end if
    
	 if CurrentPage=1 then
	    showpage TotalPut,MaxPerPage,"addgroupuser.asp"
		showContent
        showpage TotalPut,MaxPerPage,"addgroupuser.asp"
	 else 
	    if (CurrentPage-1)*MaxPerPage<TotalPut then
		 objrs.move (CurrentPage-1)*MaxPerPage
		 dim bookmark
		 bookmark=rs.bookmark
		 showpage TotalPut,MaxPerPage,"addgroupuser.asp"
		 showContent
   	    showpage TotalPut,MaxPerPage,"addgroupuser.asp"
		else
		 CurrentPage=1
         showpage TotalPut,MaxPerPage,"addgroupuser.asp"
		 showContent
	    showpage TotalPut,MaxPerPage,"addgroupuser.asp"
		end if
	 end if
	 objrs.close

sub showcontent
dim i
   i=0
%>
<center>
<table border=1 width="80%" borderColor=#ceac79  borderColorDark=#533e1e borderColorLight=#bf924d cellPadding=0  cellSpacing=1>

<tr><STRONG>
   <td></td>
  <td>姓名</td>
  <td>性别</td>
  <td>联系信息</td>
  <td>状态</td>
  <td>所属代理商</td>
  <td>版本</td>
  <td>公司名</td>
  <td>用户对象名</td>
 </tr></STRONG>
 
<% do while not objrs.eof%> 
<tr>
<td align=top ><%  my_link="addgroupuserout.asp" &"?userid="& objrs(0) %><a href="<%=my_link%>">添加</a></td>  
  <td><%=objrs(4)%> </td>
  <td><%=objrs(5)%> </td>
  <td><%=objrs(6)%> </td>
  <td><%=objrs(11)%> </td>
  <td><%=objrs(12)%> </td>
  <td><%=objrs(13)%> </td>
  <td><%=objrs(14)%> </td>
  <td><%=objrs(15)%> </td>
</tr>  
  <% 
i=i+1
if i>MaxPerPage then exit do
objrs.movenext
loop

objrs.close  
set objdml=nothing
%>
</center>
</table>
<!--
<table>
<tr>  
<td valign=top></td>  
<%  
for j=1 to 6 
%>  
    <td border=1><b>  
<%=objrs(j).name%>  
    </b></td>  
<%  
next  
for j=8 to howmanyfield
%>  
   <td border=1><b>  
<%=objrs(j).name%>  
   </b></td>  
<%  
next  
%>  
</tr> 
<tr> 
<% do while not objrs.eof %>  
<td align=top ><%  my_link="addgroupuserout.asp" &"?userid="& objrs(0) %><a href="<%=my_link%>">添加</a></td>  
<%  
for j=1 to 6 
%>  
   <td valign=top border=1>  
<%=objrs(j)%></td>  
<%  
next  
%> 
<%  
for j=8 to howmanyfield
%>  
   <td valign=top border=1>  
<%=objrs(j)%></td>  
<%  
next  
%>   
</tr> 
 <% 
i=i+1
if i>MaxPerPage then exit do
objrs.movenext
loop
objrs.close  
set objdml=nothing
%>
</table>
-->
</center>
<%end sub%>
<%

function showpage(totalnumber,maxperpage,filename)
  dim n
  if totalnumber mod maxperpage=0 then
     n= totalnumber \ maxperpage
  else
     n= totalnumber \ maxperpage+1
  end if
  response.write "<form type=Post action="&filename&">"
  if CurrentPage<2 then
  	
	response.write "<font color='999966'>首页 上一页</font>&nbsp;"
  else
  	
    response.write "<a href="&filename&"?page=1&which="&groupid&">首页</a>&nbsp;"
    response.write "<a href="&filename&"?page="&CurrentPage-1&"&which="&groupid&">上一页</a>&nbsp;"
  end if
  if n-currentpage<1 then
    response.write "<font color='999966'>下一页 尾页</font>"
  else
    response.write "<a href="&filename&"?page="&(CurrentPage+1)&"&which="&groupid&">下一页</a> <a href="&filename&"?page="&n&"&which="&groupid&">尾页</a>"
  end if
   response.write "&nbsp;页次：<strong><font color=red>"&CurrentPage&"</font>/"&n&"</strong>页 "
   response.write "&nbsp;共查到<b>"&totalnumber&"</b>个记录 <b>转到：</b>"
  response.write "<input type='text' name='page' size=4 maxlength=10 value="&currentPage&">"
   response.write "<input type='hidden' name='which' value="&groupid&">"
   response.write "<input type='submit'  value=' Goto ' name='cndok'></span></form>"
end function
%>
</div></center>
</body>
</html>

