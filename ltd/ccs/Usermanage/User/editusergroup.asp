<%@ Language=VBScript %>
<!--#include file="dbclass.asp"-->
      <%
      dim userid
      userid=Request.QueryString ("userid")
      on error resume next
      dim objF
          set objF=server.CreateObject ("Com_UserManage.ClsUserManage")
      dim UserGroup
          set UserGroup=server.CreateObject ("adodb.recordset")
          set usergroup=objF.GetAllGroup (Locale)
          if err.number<>0 then
         ierror=err.number
        err.clear
     
      response.redirect "../../Sorry.asp?Errorno="&ierror
      end if
      set objF=nothing
      %>
<form name=editusergroup method="post" action="editusergroup_result.asp?userid=<%=userid%>&b=<%=UserGroup.RecordCount%>">
<table>
    <tr> 
      <td width="15%" valign="top" bgcolor="#006666">
        <p align="right"><font color="#FFFFFF">用户所在组:</font></p>
      </td>
      <td width="85%" bgcolor="#FFFFFF" ><table> <%
      Dim Groupid
          Groupid="GroupID"
      on error resume next
      dim objB
          set ObjB=server.CreateObject ("Com_UserManage.ClsUserManage")
      dim ObjGroup
            set ObjGroup=server.CreateObject ("adodb.recordset")
      set objGroup=objB.GetUserGroup (UserID,locale)
      if err.number<>0 then
         ierror=err.number
        err.clear
      set objB=nothing
      response.redirect "../../Sorry.asp?Errorno="&ierror
      end if
      UserGroup.MoveFirst
      i=0 
      for f=0 to UserGroup.RecordCount-1%>
      	<tr align=left>
			<%for n=0 to 3%>
			  <%if i>USerGroup.RecordCount-1 then%>
				<%exit for%>
			  <%end if%>
				<td align=left width="16.7%"> 
     <% dim Grou
          Grou=GroupID&i
      %> 
        <input type=checkbox name="<%=Grou%>" value="<%=UserGroup("GroupID")%>" <%
      ObjGroup.MoveFirst 
      FOR j=1 to ObjGroup.RecordCount
       if UserGroup("GroupID")=ObjGroup("GroupID") then%>Checked<%
       end if 
       ObjGroup.MoveNEXT 
       next
       %>>
        <%=UserGroup("GroupName")%> </td><%
        i=i+1 
    UserGroup.MoveNext 
   %>
   <% next%><%f=f+n-1%></tr><%next%>
  </table></td>
    </tr>
     </table>
<table>
<tr>
<td>
<input name=submit type=submit value="提交">
</td>
</tr>
</table>
  </form>                
      <% set objF=nothing%> 