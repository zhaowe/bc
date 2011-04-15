<!--#include file="public.inc"--->
<%
dim groupid
dim dbjdml
dim objrs
dim ierrno
dim description
dim  groupuser
dim i
dim funinfo
dim groupfun

    'groupid=Request.QueryString ("which")
    groupid="{77D2D6B6-A335-49A1-BC34-6C61408FA41D}"
    description=Request.Form("description")
  
set objdml=server.createobject("com_usermanage.clsusermanage")
on  error resume next
set funinfo=server.CreateObject("adodb.recordset")
set funinfo=objdml.GetAllFunction(locale)
set groupfun=server.CreateObject("adodb.recordset")
set groupfun=objdml.GetGroupFunction(groupid,locale)
set objrs=server.CreateObject("adodb.recordset") 
set objrs=objdml.GetGroupInfo( groupid,locale)
set groupuser=server.CreateObject("adodb.recordset")
set groupuser=objdml.GetGroupUser(groupid,application("UseObject"))
if Err.number<>0 then
	iErrNo = Err.number
	Err.Clear
	Response.Redirect "../../Sorry.asp?ErrorNo=" & iErrNo
set objdml=nothing	
End If
    session("groupid")=groupid
    howmanyfields=groupuser.Fields.Count-1
checked="checked"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel=stylesheet href="style.css">
<title>修改组信息</title>
</head>


<body background="images/bg.gif" style="font-size:10.5pt">
<div align=center style="width: 575; height: 43"> 
  <p><b><font color="blue" face="楷体_GB2312" size="6">修改组信息</font></b> <br>     
         
 <hr>
 
 <form name=editgroup action="editgroupout.asp" method="post">   
 <div align=center style="width: 641; height: 235">                                          
<table border="0" width="513">                                                 
  <tr>                                 
     <td width="130" align="right">组名:</td>                                 
     <td width="369"><input type=test name="groupname" value="<% =objrs("groupname")%>"></td>                                                               
  </tr>
  <tr>                                 
    <td width="130" align="right">语言版本号:</td>                                 
    <td width="369"><input type=test name="locale" value="<% =objrs("locale")%>"></td>                                                               
  </tr>
  <tr>                                 
    <td width="130" align="right">描述:</td>                                 
    <td width="369"><input type=test name="description" value="<% =objrs("description")%>" maxlength=50></td>                                                               
  </tr>                                   
</table>                                                      
<div align=right>             
  <input type=submit name="submit_button" value=提交>             
  <input type=reset name="reset_button" value=重置>                 
 </form>  
 <hr> 
 

<form type=post name=editgroupfun action="editgroupfun.asp">
  <TABLE WIDTH=75% BORDER=1 CELLSPACING=1 CELLPADDING=1>
  <TR> <td>所有的功能:</td><% 
  dim  functionid
  functionid="Functionid"
   
        
         %>
		<%
		Funinfo.movefirst
		for i=0 to funinfo.RecordCount-1 
        dim  func
        func=functionid&i
        %>
		
	<td>
	 <input type=checkbox name="<%Func%>" value="<%=funinfo("FunctionId")%>" <% for j=0 to groupfun.RecordCount-1%><%if funinfo("functionid")=Groupfun("functionid") then%> checked <%end if%>
	<% 
	
       groupfun.movenext
      next
	%>>	<%=funinfo("FunctionName")%>	
    </td>
		<% funinfo.movenext
		next
		  
           
		%>

		
	</TR>
	<TR>
		<TD><input type="submit" name="submit" value="提交"><input type="reset" name="button" value="放弃"></TD>
	
	</TR>
	
</TABLE>
</form>

 <%
 mypage=Request.QueryString("whichpage")
 if mypage="" then
 mypage=1
 end if
 mypagesize=Request.QueryString("pagesize")
 if mypagesize="" then
 mypagesize=10 
 end if
 groupuser.CursorLocation=aduseclient
 groupuser.CacheSize=5
 groupuser.movefirst
 groupuser.PageSize=mypagesize
 maxcount=cint(groupuser.PageCount)
 groupuser.AbsolutePage=mypage
 Response.Write"page"& mypage &"of"&maxcount&"<br>"
 
 
 %>
  

<div align=center>在组中的所有用户</div>
<table border=1 align=center>  
<tr>  
<td valign=top></td> 
<td> </t>
<%  
for i=5 to 7 
%>  
   <td>  
<%=groupuser(i).name%>  
   </td>  
<%  
next  
%>
<%  
for i=9 to howmanyfields
%>  
   <td>  
<%=groupuser(i).name%>  
   </td>  
<%  
next  
%>
</tr>
 
   <% 
do while not groupuser.eof
 %>  
  <tr>  
   <td align=top ><%my_link="delgroupuser.asp" &"?userid="& groupuser(0)%>  
   <a href="<%=my_link%>">删除</a>
   </td>    
   <td align=top >  

<%  
for i=5 to 7
%>  
   <td valign=top >  
<%=groupuser(i)%></td>  
<%  
next  
%> 
<%  
for i=9 to howmanyfields
%>  
   <td valign=top >  
<%=groupuser(i)%></td>  
<%  
next  
%> 
</tr>  
<%  
groupuser.movenext  
loop  
groupuser.close  
set objdml=nothing
%> 
</table>

<%pad="0"
scriptname=Request.ServerVariables("script_name")
for counter=1 to maxcount
 if counter>=10 then
 pad=""
 end if
 ref="<a href='" &scriptname & "?whichpage="& counter
 ref=ref&"&pagesize="& mypagesize&"'>"&pad &counter&"</a>"
 Response.Write ref&""
 if counter mod 10=0 then
 Response.Write"<br>"
 end if
 next
 
%>
    
</body>      
  
</html>                      





 