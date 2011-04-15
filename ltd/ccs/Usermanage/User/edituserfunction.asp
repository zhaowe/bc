<%@ Language=VBScript %>
<!--#include file="dbclass.asp"-->
<%
dim userid
userid=Request.QueryString ("Userid")
      dim ObjH
          set objH=server.CreateObject ("Com_UserManage.ClsFunction")
      dim UserFuncton
          set userfunction=server.CreateObject ("adodb.recordset")
          set userfunction=ObjH.GetAllFunction (Locale)
       if err.number<>0 then
       ierror=err.number
      err.clear
      set objH=nothing
      response.redirect "../../Sorry.asp?Errorno="&ierror
      end if   
      set OBJh=nothing
      %>
<form name=edituserfunction Method="post" action="edituserfunction_result.asp?userid=<%=userid%>&a=<%=Userfunction.RecordCount%>">
<table border="0" width="730" bgcolor="#003333" cellspacing="1" cellpadding="4">
    <tr> 
      <td width="6%" valign="top" bgcolor="#006666"> 
        <p align="right"><font color="#FFFFFF">用户权限:</font></p>
      </td>
      <td width="94%" bgcolor="#FFFFFF" ><table> <%
      dim Functionid
          Functionid="Functionid"
      On Error resume next
      dim objA
          set ObjA=server.CreateObject ("Com_UserManage.ClsUserManage")
      dim ObjFunction 
     ' set ObjFunction=server.CreateObject ("adodb.recordset")
      'set objFunction=objA.GetUserFunction (UserID,Locale)
      objFunction=objA.GetFuncStr(UserID)
         'Response.Write objFunction
        ' Response.End 

      if err.number<>0 then
       ierror=err.number
      err.clear
      set objA=nothing
      response.redirect "../../Sorry.asp?Errorno="&ierror
      end if
      USERFunction.Movefirst
      m=0
      'Response.Write ","&Trim(USERFunction("ordernum")) & ","
       'L=","&Trim(USERFunction("ordernum")) & ","
      ' Response.Write L
       
     'Response.End 
     'count=1      
		for n=0 to USERFunction.RecordCount-1 %>
		<tr align=left>
			<%for i=0 to 2%>
			<%if m>USerFunction.RecordCount-1 then%>
				<%exit for%>
			  <%end if%>
				<td align=left width="16.7%">
     <%dim func 
       func=functionid&m  %> 
        <input type=checkbox name="<%=Func%>" value="<%=USERFunction("FunctionId")%>" 
	<%
       ' ObjFunction.MoveFirst 
        'Response.Write M
        'Response.end
         'for j=1 to ObjFunction.RecordCount 
        L=","&Trim(USERFunction("ordernum")) & ","
      if instr(1,OBJFunction,L)<>0 then%>checked<%end if
      %>>
        <%=USERFunction("FunctionName")%> </td><%m=m+1%><%USERFunction.MoveNext%>
     <%next%>
    <% n=n+I-1%>
		</tr>
       <%next%>
       </table>
       <table>
       <tr>
       <td>
       <input type=submit name=submit value="提交">
       </td>
       </tr>
       </table>
       </form>
  <%
  set objH=nothing
  %>