 <!--#include file="public.inc"--->
<%
dim groupid
dim dbjdml
dim objrs
dim ierrno
dim description
dim groupuser
dim j
dim funinfo
dim groupfun
groupid=Request.QueryString("which")
description=Request.Form("description")
session("groupid")=groupid
groupid=session("groupid")
on error resume next
set objdml=server.createobject("Com_UserManage1.clsUserManage1")
set obj=server.createobject("com_usermanage.clsFunction")
set funinfo=server.CreateObject("ADODB.Recordset")
set funinfo=obj.GetAllFunction(locale)
set groupfun=server.CreateObject("adodb.recordset")
set groupfun=objdml.GetGroupFunction(groupid,locale)
set objrs=server.CreateObject("adodb.recordset")
set objrs=objdml.GetGroupInfo(groupid,locale) 
set groupuser=server.CreateObject("adodb.recordset")
set groupuser=objdml.GetGroupUser(groupid,application("UseObject"))
if Err.number<>0 then
	iErrNo = Err.number
	Err.Clear
	Response.Redirect "../../Sorry.asp?ErrorNo=" & iErrNo
set objdml=nothing	
End If
set objdml=nothing
howmanyfields=groupuser.Fields.Count-1
checked="checked"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�޸�����Ϣ</title>
<link rel="stylesheet" href="../../style.css">
</head>
<script language="vbscript" >     
<!--     
sub datacheck()
 if editgroup.groupname.value="" then     
    msgbox "����������������Ϊ�� ",64,"��ע��!"      
	exit sub
 End if

 editgroup.submit     
 end sub      
-->     
</script>
<body bgcolor="#FFFFFF" style="font-size:10.5pt">
<b><font color="blue">�޸�����Ϣ</font></b> 
  <form name="editgroup" action="editgroupfun.asp" method="post">   
 <input type=hidden name="groupid" value="<% =objrs("groupid")%>">
   <input type=hidden name="ffunctionid" value="<%=objrs("ffunctionid")%>">
  <table border="0" width="640" cellpadding="4" cellspacing="1" bgcolor="#000000">
    <tr> 
      <td width="84" align="right" bgcolor="#003333"><font color="#FFFFFF">����:</font></td>
      <td width="152" bgcolor="#FFFFFF"> 
        <input type=test name="groupname" value="<% =objrs("groupname")%>">
      </td>
      <td width="40" align="right" bgcolor="#003333"><font color="#FFFFFF">����:</font></td>
      <td width="293" bgcolor="#FFFFFF"><% =objrs("description")%></td>
    </tr>
    <TR> 
      <td width="84" bgcolor="#003333" valign="top" align="right"><font color="#FFFFFF">���еĹ���:</font></td>
      <td colspan="3" bgcolor="#FFFFFF"><table align=left><td align=left width="16.7%">  
       <%session("count")=funinfo.RecordCount
       dim Functionid
          Functionid="functionid"   
		  Funinfo.movefirst
		dim i
		i=0  
       for m=0 to Funinfo.RecordCount-1%>
		  <tr align=left>
			<%for n=0 to 3%>
			  <% if i > Funinfo.RecordCount-1 then%>
				<%exit for%>
			  <%end if%>
				<td align=left width="16.7%">
				<%dim func
				func=functionid&i
				groupfun.MoveFirst%> 
				<input type=checkbox name="<%=Func%>" value="<%=funinfo("functionid")%>" 
				<%for j=0 to groupfun.RecordCount-1%><%if funinfo("functionid")=Groupfun("functionid") then%> checked <%end if%>
					<%groupfun.movenext
				next%>>
				<%=funinfo("FunctionName")%>
				</td>
		        <%funinfo.movenext%>
		        <%'n=n+1
		        i=i+1%>
			<%next%>
			<%m=m+n-1%>
		</tr>
	   <%next%>
	 
	</table></td>
    </TR>
  </TABLE>
    <input type="button" name="button" value="�ύ" onClick="datacheck">
    <input type="reset" name="button" value="����">
</form>  
<STRONG><font color="blue">�����е������û�</font></STRONG>
<div align=right> </div>
  <%
  const MaxPerPage=4
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
   
     TotalPut=groupuser.recordcount
     if TotalPut=0 then
     currentPage=0
     end if 
	 
	 if CurrentPage<1 then
	   CurrentPage=0
	 end if
	 
	 if (CurrentPage-1)*MaxPerPage>TotalPut then
	   if (TotalPut mod MaxPerPage)=0 then
	     CurrentPage=TotalPut \ MaxPerPage
		else
		 CurrentPage=TotalPut \ MaxPerPage + 1
		end if
     end if
    
	 if CurrentPage=1 then
	    showpage TotalPut,MaxPerPage,"editgroup.asp"
		showContent
        'showpage TotalPut,MaxPerPage,"editgroup.asp"
	 else 
	    if (CurrentPage-1)*MaxPerPage<TotalPut then
		 groupuser.move (CurrentPage-1)*MaxPerPage
		 dim bookmark
		 bookmark=rs.bookmark
		 showpage TotalPut,MaxPerPage,"editgroup.asp"
		 showContent
   	    'showpage TotalPut,MaxPerPage,"editgroup.asp"
		else
		 CurrentPage=1
         showpage TotalPut,MaxPerPage,"editgroup.asp"
		 showContent
	    'showpage TotalPut,MaxPerPage,"editgroup.asp"
		end if
	 end if
	 groupuser.close

sub showcontent
dim i
   i=0
%> 
  <FORM action="DelGroupUser.asp" method="get" name="form1">
    
  <table border=0 cellPadding=4  cellSpacing=1 width="610" bgcolor="#000000">
    <tr bgcolor="#003333"> <font color="#FFFFFF"></font>
        
      <td width="43"><font color="#FFFFFF">ѡ��</font></td>
        
      <td width="43"><font color="#FFFFFF">����</font></td>
        
      <td width="43"><font color="#FFFFFF">�Ա�</font></td>
        
      <td width="82"><font color="#FFFFFF">��ϵ��Ϣ</font></td>
        
      <td width="43"><font color="#FFFFFF">״̬</font></td>
        
      <td width="99"><font color="#FFFFFF">����������</font></td>
        
      <td width="64"><font color="#FFFFFF">��˾��</font></td>
      <td width="64"><font color="#FFFFFF">����ʱ��</font></td>
    </tr>
      <% do while not groupuser.eof%> 
      
    <tr bgcolor="#FFFFFF"> 
      <td align=top width="43" > 
        <INPUT id=radio1 name=userid type=radio value="<%=groupuser(0)%>" <%if i mod 5 =1 then Response.Write "checked" end if %>>
        </td>
        
      <td width="43"><%=groupuser(5)%> </td>
        
      <td width="43"><%=groupuser(6)%> </td>
        
      <td width="82"><%=groupuser(7)%> </td>
      <td width="43"><%f=trim(groupuser(11))%><% if f="E" then%> ��Ч <%else%>��ͣ <%end if%> </td>
        
      <td width="99"><%=groupuser(12)%> </td>
        
      <td width="64"><%=groupuser(14)%> </td>
      <td width="64"><%=groupuser("enddate")%> </td>
        <!--
  <td><%=groupuser(15)%> </td>--> </tr>
      <% 
i=i+1
if i>MaxPerPage then exit do
groupuser.movenext
loop
groupuser.close  
set objdml=nothing
%> 
    </table>
    
  <table width="610" border="0" cellspacing="1" cellpadding="4">
    <tr>
        
      <td width="314"> 
        <input name=button2 type=submit value=ɾ�� >
        <input type="reset" name="button3" value="����" onClick="self.history.back()">
      </td>
        
      <td width="275" align="center">[ <a href="groupinfo.asp">���������ҳ</a> ] [ 
        <a href=<%="addgroupuser.asp"&"?which="& groupid%>>�����û�</a> ] [ <a href=<%="gsontablerun.asp"&"?which="& groupid%>>�ӱ����</a> 
        ]</td>
      </tr>
    </table>
  </FORM>

<%end sub
function showpage(totalnumber,maxperpage,filename)
  dim n
  if totalnumber mod maxperpage=0 then
     n= totalnumber \ maxperpage
  else
     n= totalnumber \ maxperpage+1
  end if
  response.write "<form type=Post action="&filename&" id=form1 name=form1>"
  if CurrentPage<2 then  	
	response.write "<font color='999966'>��ҳ ��һҳ</font>&nbsp;"
  else   	
    response.write "<a href="&filename&"?page=1&which="&groupid&">��ҳ</a>&nbsp;"
    response.write "<a href="&filename&"?page="&CurrentPage-1&"&which="&groupid&">��һҳ</a>&nbsp;"
  end if
  if n-currentpage<1 then
    response.write "<font color='999966'>��һҳ βҳ</font>"
  else
    response.write "<a href="&filename&"?page="&(CurrentPage+1)&"&which="&groupid
    response.write ">��һҳ</a> <a href="&filename&"?page="&n&"&which="&groupid&">βҳ</a>"
  end if
   response.write "&nbsp;ҳ�Σ�<strong><font color=red>"&CurrentPage&"</font>/"&n&"</strong>ҳ "
   response.write "<input type='hidden' name='which' value="&groupid&">"
   response.write "</span></form>"
end function
%> 
</body>
</html>


 