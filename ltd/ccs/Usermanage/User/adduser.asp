<!--#include file="dbclass.asp"-->
<%
'if session("loginid")="" then
'Response.Redirect "login.htm"
%>
<script ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
sub btnQuery_onclick
dim name
dim loginid
dim contactinfo
dim password
cnstr=""
titlestr="[您请注意]"+chr(13)+chr(13) 
errstr=""
tempname=trim(document.userinfo.UserName.value)
if(tempname="")then
cnt=cnt+1
cntstr=cstr(cnt)
errstr=errstr+cntstr+"."+"请输入您的姓名"+chr(13)
else
document.userinfo.username.value=tempname
end if
temploginid=trim(document.userinfo.loginid.value)
if(temploginid="")then
cnt=cnt+1
cntstr=cstr(cnt)
errstr=errstr+cntstr+"."+"请输入您的注册名"+chr(13)
else
document.userinfo.LoginID.value=temploginid
end if
tempcontactinfo=trim(document.userinfo.contactinfo.value)
if(tempcontactinfo="")then
cnt=cnt+1
cntstr=cstr(cnt)
errstr=errstr+cntstr+"."+"请输入您的联系信息"+chr(13)
else
document.userinfo.contactinfo.value=tempcontactinfo
end if
temppassword=trim(document.userinfo.password.value)
if(temppassword="")then
cnt=cnt+1
cntstr=cstr(cnt)
errstr=errstr+cntstr+"."+"请输入您的密码"+chr(13)
else
document.userinfo.password.value=temppassword
end if
if len(temppassword)>=10 then
cnt=cnt+1
cntstr=cstr(cnt)
errstr=errstr+cntstr+"."+"请您输入一个小于10位的密码"+chr(13)
end if
enddate1=userinfo.EndDate1.value
enddate2=userinfo.EndDate2.value
enddate3=userinfo.EndDate3.value
enddate=enddate1 &"-"& enddate2 &"-"& enddate3
if isdate(enddate) then
enddate=cdate(enddate)
else
cnt=cnt+1
cntstr=cstr(cnt)
errstr=errstr+cntstr+"."+"结束时间不是一个正确的时间值"+chr(13)
end if
dim nowtime
nowtime=Dateadd("D",6,now())
if enddate < nowtime then
cnt=cnt+1
cntstr=cstr(cnt)
errstr=errstr+cntstr+"."+"您输入的时间必须晚于当前时间七天！"+chr(13)
end if
if cnt<>0 then
alert(errstr)
else
userinfo.submit()
end if
end sub

-->
</script>
 <%
      On Error resume Next
      dim objD
          set ObjD=server.CreateObject ("Com_UserManage.ClsUserManage")
      dim objGroup
          set ObjGroup=Server.CreateObject ("adodb.recordset")
          set ObjGroup=objD.GetAllGroup (locale)
           
          if err.number<>0 then
             ierror=err.number
             err.clear
             set objD=nothing
             response.redirect "../../Sorry.asp?Errorno="&Ierror
          end if
          b=objGroup.RecordCount%>
<%
      dim objC
          set ObjC=server.CreateObject ("Com_UserManage.ClsFunction")
      dim objFunction
          set ObjFunction=Server.CreateObject ("adodb.recordset")
          set ObjFunction=ObjC.GetAllFunction (locale)
          if err.number<>0 then
             ierror=err.number
             err.clear
             set objC=nothing
             response.redirect  "../../Sorry.asp?Errorno="&Ierror
          end if
          a=objFunction.RecordCount%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>增加用户</title>
<link rel="stylesheet" href="../../style.css">
</head>
<body bgcolor="#FFFFFF">
<b>增加用户</b> <br>
<form name="userinfo" Method="Post" action="adduser_result.asp?a=<%=objFunction.RecordCount %>&amp;b=<%=objGroup.RecordCount %>">
    
  <table width="610" border="0" cellspacing="1" cellpadding="4" bgcolor="#000000">
    <tr> 
      <td  align="right" bgcolor="#003333" width="16%"><font color="#FFFFFF">用户姓名:</font></td>
      <td bgcolor="#FFFFFF" width="36%" > 
        <input type="text" name="UserName" maxlength="50">
      </td>
      <td  align="right" bgcolor="#003333" width="15%"><font color="#FFFFFF">用户性别:</font></td>
      <td bgcolor="#FFFFFF" width="33%" > 
        <input type="radio" name="sex" value="M" checked>
        男&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
        <input type="radio" name="sex" value="F">
        女</td>
    </tr>
    <tr> 
      <td  align="right" bgcolor="#003333" width="16%"><font color="#FFFFFF">用户ID:</font></td>
      <td bgcolor="#FFFFFF" width="36%" > 
        <input type="text" name="loginid" maxlength="20">
      </td>
      <td  align="right" bgcolor="#003333" width="15%"><font color="#FFFFFF">部门名称:</font></td>
      <td bgcolor="#FFFFFF" width="33%" > 
        <select name="companyid" size="1">
          <% 
      On Error resume Next 
      dim obj 
          set Obj=server.CreateObject("Com_UserManage.ClsUserManage") 
      dim objCompany 
          set ObjCompany=Server.CreateObject ("adodb.recordset") 
          set ObjCompany=Obj.GetCompany(locale) 
          if err.number<>0 then 
             ierror=err.number 
             err.clear 
             set obj=nothing 
             response.redirect "../../Sorry.asp?Errorno="&Ierror 
          end if 
          objCompany.MoveFirst  
          do while not objCompany.EOF  
           
       %> 
          <option value="<%=ObjCompany("CompanyID")%>"><%=ObjCompany("CompanyName")%></option>
          <% 
       objCompany.MoveNext  
       loop 
       set obj=nothing 
       %> 
        </select>
      </td>
    </tr>
    <tr> 
      <td  align="right" height="35" bgcolor="#003333" width="16%"><font color="#FFFFFF">用户密码:</font></td>
      <td  height="35" bgcolor="#FFFFFF" width="36%" > 
        <input type="Password" name="password" maxlength="10">
      </td>
      <td  align="right" height="35" bgcolor="#003333" width="15%"><font color="#FFFFFF">类别:</font></td>
      <td  height="35" bgcolor="#FFFFFF" width="33%"> 
        <select name="agentID" size="1">
          <%
      On Error resume Next
      dim objA
          set ObjA=server.CreateObject ("Com_UserManage.ClsUserManage")
      dim objAgent
          set ObjAgent=Server.CreateObject("adodb.recordset")
          set ObjAgent=ObjA.GetAgent()
          if err.number<>0 then
             ierror=err.number
             err.clear
             set objA=nothing
             response.redirect "../../Sorry.asp?Errorno="&Ierror
          end if
          objAgent.MoveFirst 
          do while not objAgent.EOF 
          
       %> 
          <option value="<%=ObjAgent("AgentID")%>"><%=ObjAgent("AgentName")%></option>
          <%
       objAgent.MoveNext 
       loop
       set objA=nothing
       %> 
        </select>
      </td>
    </tr>
    <tr> 
      <td  align="right" bgcolor="#003333" width="16%"><font color="#FFFFFF">联系方式:</font></td>
      <td bgcolor="#FFFFFF" width="36%" > 
        <input type="text" name="contactinfo" maxlength="50">
      </td>
      <td bgcolor="#003333" width="15%"><font color="#FFFFFF"></font></td>
      <td bgcolor="#FFFFFF" width="33%"></td>
    </tr>
    <tr> 
      <td  align="right" bgcolor="#003333" width="16%"><font color="#FFFFFF">结束时间:</font></td>
      <td bgcolor="#FFFFFF" width="36%" > 
        <p> 
          <select name="EndDate1">
            <% For i=2010 to 2050 
            
            %> 
            <option value="<%=i%>" <%if i=Year(Thistime) then%>selected<%end if%>><%=I%></option>
            <%next%> 
          </select>
          年 
          <select name="EndDate2">
            <% For h=1 to 12 
            
            %> 
            <option value="<%=h%>" <%if h=Month(Thistime) then%>selected<%end if%>><%=h%></option>
            <%next%> 
          </select>
          月 
          <select name="EndDate3">
            <% For j=1 to 31 
            
            %> 
            <option value="<%=j%>" <%if j=Day(Thistime) then%>selected<%end if%>><%=j%></option>
            <%next%> 
          </select>
          日 
      </td>
      <td bgcolor="#003333" width="15%" ><font color="#FFFFFF"></font></td>
      <td bgcolor="#FFFFFF" width="33%" >&nbsp;</td>
    </tr>
    <tr> 
      <td  align="right" bgcolor="#003333" width="16%"><font color="#FFFFFF">员工号:</font></td>
      <td bgcolor="#FFFFFF" width="36%" > 
       <input type="text" name="no" maxlength="16">
      </td>
      <td bgcolor="#003333" width="15%" ><font color="#FFFFFF"></font></td>
      <td bgcolor="#FFFFFF" width="33%" >&nbsp;</td>
    </tr>
  </table> 
<br>
 <input type="button" name="btnQuery" id=btnquery value="提交" >               
  <input type="reset" name="reset" value="重置">
  <input type="button" value="返回" onclick="self.history.back()">
  <% 
 objFunction.Close 
 objGroup.Close 
 objAgent.Close 
 ObjCompany.close 
 objUserObject.Close  
 set ObjFunction=nothing 
 set ObjAgent=nothing 
 set objcompany=nothing 
 set objGroup=nothing 
 %> 
</form>                           
</body>                                                            
                                                            
</html>                                                            
 
