<!--#include file="public.inc"--->
<html>
<head>
<title>�����</title>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Sub btnQuery_onclick
	frminfo.action = "groupinfo.asp?tiaojian=" & frminfo.tiaojian.value & "&textfield=" & frminfo.textfield.value
	frminfo.submit  
End Sub
-->
</SCRIPT> 
<link rel="stylesheet" href="../../style.css">
</head> 
<body bgcolor="#FFFFFF">
<font size=5 color=blue><strong>�����</strong></font>
<form name="frminfo" method="post" action="groupinfo.asp" >
  <table width="610" border="0" cellspacing="1" cellpadding="4" height="15" bgcolor="#006699" >
      <tr> 
        <td colspan="2"><font color="#FFFFFF">��ѯ������</font> 
          <select name="tiaojian">
            <option value="groupname">����</option>
            <option value="description">����</option>
            <option value="locale">�汾</option>
          </select>
          <input type="text" name="textfield">
          <input type="button" name="BtnQuery" id=BtnQuery  value="��ʼ��ѯ">
          <input type="reset" name="reset" value="ȡ��">
        </td>
      </tr>
    </table>
   
        </form>     
<%
  On Error resume next
dim tiaojian
dim textfield
dim obj
dim groupinfo
    tiaojian=Request.QueryString ("tiaojian")
    textfield=Request.QueryString ("textfield")
    textfield=trim(textfield)
 on error resume next   
set obj=Server.CreateObject("Com_UserManage1.clsUserManage1")
set GroupInfo=Server.CreateObject ("ADODB.Recordset")
set Groupinfo=obj.SearchGroupInfo(textfield,tiaojian)
if err.number<>0 then
   Ierror=err.number
   err.clear
   set Objdml=nothing
   'Response.End 
   response.redirect "../../Sorry.asp?Errorno=" &ierror
end if 

howmanyfield=groupinfo.fields.count-1
  %>
  
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
   
     TotalPut=groupinfo.recordcount
     if TotalPut=0 then
     CurrentPage=0
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
	    showpage TotalPut,MaxPerPage,"groupinfo.asp"
		showContent
	 else 
	    if (CurrentPage-1)*MaxPerPage<TotalPut then
		 groupinfo.move (CurrentPage-1)*MaxPerPage
		 dim bookmark
		 bookmark=rs.bookmark
		 showpage TotalPut,MaxPerPage,"groupinfo.asp"
		 showContent
		else
		 CurrentPage=1
         showpage TotalPut,MaxPerPage,"groupinfo.asp"
		 showContent
		end if
	 end if
	 groupinfo.close

sub showcontent
dim i
   i=1
%> 
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function button1_onclick() {
	form1.text.value="edit"
	form1.submit()
	return true 
}

function button2_onclick() {
	form1.text.value="del"
	form1.submit()
	return true 
}

//-->
</SCRIPT>
  <%  if not GroupInfo.EOF then %> 
<FORM action="bosom.asp" method="post" name="form1">

  <table border=0 cellPadding=4  cellSpacing=1 width="610" bgcolor="#000000">
    <STRONG> </STRONG> 
    <tr bgcolor=#003333> 
      <td valign=top width="42"> 
        <P><strong><font color="#FFFFFF">����</font></strong></P>
      </td>
      <td valign=top width="279"> 
        <P><strong><font color="#FFFFFF">����</font></strong></P>
      </td><!--
      <td valign=top width="105"> 
        <P><strong><font color="#FFFFFF">�汾</font></strong></P>
      </td>-->
      <td valign=top width="143"> 
        <P><strong><font color="#FFFFFF">����</font></strong></P>
      </td>
    </tr>
    <STRONG> </STRONG> 
    <tr> <% do while not GroupInfo.eof%> 
      <td align=top bgcolor="#FFFFFF" width="42" > 
        <INPUT id=radio1 name=groupid type=radio value="<%=GroupInfo(0)%>" <%if i mod 10 =1 then Response.Write  "checked" end if%>>
      </td>
      <td  valign=top border="1" bgcolor="#FFFFFF"><%=groupinfo("groupname")%></td>
      <td  valign=top border="1" bgcolor="#FFFFFF"><%=groupinfo("description")%></td>
      </tr>
    <% 
i=i+1
if i>MaxPerPage then exit do  
GroupInfo.movenext
loop
GroupInfo.close  
set Objdml=nothing
%> 
  </TABLE>
  <%else  Response.Write "��ǰû�м�¼" end if %>
  <table width="610" border="0" cellspacing="1" cellpadding="4">
    <tr>
      <td width="424"> 
        <input name=text type=hidden>
        <input name=button1 type=button value=�޸� language=javascript onClick="return button1_onclick()">
        <input name=button2 type=button value=ɾ�� language=javascript onClick="return button2_onclick()">
      </td>
      <td width="165">��Ҫ��ز���:[ <a href="addgroup.asp">������</a> ]</td>
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
  response.write "<form method=Post action="&filename&">"
  if CurrentPage<2 then
  	
  else
   	
    response.write "<a href="&filename&"?page=1>��ҳ</a>&nbsp;"
    response.write "<a href="&filename&"?page="&CurrentPage-1&">��һҳ</a>&nbsp;"
  end if
  if n-currentpage<1 then
  else
    response.write "<a href="&filename&"?page="&(CurrentPage+1)&">��һҳ</a> <a href="&filename&"?page="&n&">βҳ</a>"
  end if
   response.write "&nbsp;��<strong><font color=red>"&CurrentPage&"</font>/"&n&"</strong>ҳ "

  
   response.write "</span></form>"
end function
%>

</body>
</html>


