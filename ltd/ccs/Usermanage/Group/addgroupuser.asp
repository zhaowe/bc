<!--#include file="public.inc"-->
<%
dim groupinfo
dim groupid
dim objdml
dim groupuser
dim howmanyfield
   groupid=Request.QueryString("which")
  session("groupid")=groupid
set objdml=server.CreateObject ("com_usermanage.clsusermanage") 
on error resume next
set groupinfo=server.CreateObject("adodb.recordset")
set groupinfo=objdml.GetGroupInfo(groupid)
set objrs=server.createobject("adodb.recordset")
set objrs=objdml.GetAllUser(locale,UseObject)
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
<title>增加组用户管理</title>
<link rel="stylesheet" href="../../style.css">
</head>

<body bgcolor="#FFFFFF">
<p><font color=blue><strong>增加组用户</strong></font> </p>
<TABLE WIDTH=610 cellspacing="1" cellpadding="4" bgcolor="#000000"  >
    <TR>
		
      <TD width="10%" bgcolor="#003333"><font color="#FFFFFF">组名:</font></TD>
      <TD width="20%" bgcolor="#FFFFFF"><%=groupinfo(3)%></TD>
		
      <TD width="10%" bgcolor="#003333"><font color="#FFFFFF">描述:</font></TD>
      <TD bgcolor="#FFFFFF"><%=groupinfo(1)%></TD>
	</TR>
</TABLE>
<%
  const MaxPerPage=8
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
        'showpage TotalPut,MaxPerPage,"addgroupuser.asp"
	 else 
	    if (CurrentPage-1)*MaxPerPage<TotalPut then
		 objrs.move (CurrentPage-1)*MaxPerPage
		 dim bookmark
		 bookmark=rs.bookmark
		 showpage TotalPut,MaxPerPage,"addgroupuser.asp"
		 showContent
   	    'showpage TotalPut,MaxPerPage,"addgroupuser.asp"
		else
		 CurrentPage=1
         showpage TotalPut,MaxPerPage,"addgroupuser.asp"
		 showContent
	    'showpage TotalPut,MaxPerPage,"addgroupuser.asp"
		end if
	 end if
	 objrs.close

sub showcontent
dim i
   i=1
%>
<FORM action="addgroupuserout.asp" method="get" name="form1">  

    
  <table border=0 width="610" cellPadding=4  cellSpacing=1 bgcolor="#000000">
    <tr bgcolor="#003333"> 
      <td><font color="#FFFFFF"></font></td>
      <td><font color="#FFFFFF">姓名</font></td>
      <td><font color="#FFFFFF">性别</font></td>
      <td><font color="#FFFFFF">联系信息</font></td>
      <td><font color="#FFFFFF">状态</font></td>
      <td><font color="#FFFFFF">所属代理商</font></td>
      <td><font color="#FFFFFF">结束时间</font></td>
      <td><font color="#FFFFFF">公司名</font></td>
 <% do while not objrs.eof %>     <tr bgcolor="#FFFFFF">
      <td align=top > 
        <INPUT id=radio1 name=userid type=radio value="<%=objrs(0)%>" <%if i mod 8 =1 then Response.Write "checked" end if %>>
      </td>
      <td><%=objrs("name")%></td>
      <td><%=objrs("sex")%></td>
      <td><%=objrs("contactinfo")%></td>
      <td><%f=trim(objrs("status"))%><% if f="E" then%> 有效 <%else%>暂停 <%end if%> </td>
      <td><%=objrs("agentname")%></td>
      <td><%=objrs("enddate")%></td>
      <td><%=objrs("companyname")%></td>
</tr>     
    <% 
i=i+1
if i>MaxPerPage then exit do
objrs.movenext
loop
objrs.close  
set objdml=nothing
%> 
  </table><br>
<INPUT name=button2 type=submit value=增加 >
  <input name=button22 type=button value=返回 onclick="self.history.back()">
</FORM>
  
<%end sub%> <%

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
  ' response.write "&nbsp;共查到<b>"&totalnumber&"</b>个记录 "
   response.write "<input type='hidden' name='which' value="&groupid&">"
   response.write "</span></form>"
end function
%> 
</body>
</html>

