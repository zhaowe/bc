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
titlestr="[����ע��]"+chr(13)+chr(13) 
errstr=""
tempname=trim(document.userinfo.UserName.value)
if(tempname="")then
cnt=cnt+1
cntstr=cstr(cnt)
errstr=errstr+cntstr+"."+"��������������"+chr(13)
else
document.userinfo.username.value=tempname
end if
temploginid=trim(document.userinfo.loginid.value)
if(temploginid="")then
cnt=cnt+1
cntstr=cstr(cnt)
errstr=errstr+cntstr+"."+"����������ע����"+chr(13)
else
document.userinfo.LoginID.value=temploginid
end if
tempcontactinfo=trim(document.userinfo.contactinfo.value)
if(tempcontactinfo="")then
cnt=cnt+1
cntstr=cstr(cnt)
errstr=errstr+cntstr+"."+"������������ϵ��Ϣ"+chr(13)
else
document.userinfo.contactinfo.value=tempcontactinfo
end if
temppassword=trim(document.userinfo.password.value)
if(temppassword="")then
cnt=cnt+1
cntstr=cstr(cnt)
errstr=errstr+cntstr+"."+"��������������"+chr(13)
else
document.userinfo.password.value=temppassword
end if
if len(temppassword)>=10 then
cnt=cnt+1
cntstr=cstr(cnt)
errstr=errstr+cntstr+"."+"��������һ��С��10λ������"+chr(13)
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
errstr=errstr+cntstr+"."+"����ʱ�䲻��һ����ȷ��ʱ��ֵ"+chr(13)
end if
dim nowtime
nowtime=Dateadd("D",6,now())
if enddate < nowtime then
cnt=cnt+1
cntstr=cstr(cnt)
errstr=errstr+cntstr+"."+"�������ʱ��������ڵ�ǰʱ�����죡"+chr(13)
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
<title>�����û�</title>
<link rel="stylesheet" href="../../style.css">
</head>
<body bgcolor="#FFFFFF">
<b>�����û�</b> <br>
<form name="userinfo" Method="Post" action="adduser_result.asp?a=<%=objFunction.RecordCount %>&amp;b=<%=objGroup.RecordCount %>">
    
  <table width="610" border="0" cellspacing="1" cellpadding="4" bgcolor="#000000">
    <tr> 
      <td  align="right" bgcolor="#003333" width="16%"><font color="#FFFFFF">�û�����:</font></td>
      <td bgcolor="#FFFFFF" width="36%" > 
        <input type="text" name="UserName" maxlength="50">
      </td>
      <td  align="right" bgcolor="#003333" width="15%"><font color="#FFFFFF">�û��Ա�:</font></td>
      <td bgcolor="#FFFFFF" width="33%" > 
        <input type="radio" name="sex" value="M" checked>
        ��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
        <input type="radio" name="sex" value="F">
        Ů</td>
    </tr>
    <tr> 
      <td  align="right" bgcolor="#003333" width="16%"><font color="#FFFFFF">�û�ID:</font></td>
      <td bgcolor="#FFFFFF" width="36%" > 
        <input type="text" name="loginid" maxlength="20">
      </td>
      <td  align="right" bgcolor="#003333" width="15%"><font color="#FFFFFF">��������:</font></td>
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
      <td  align="right" height="35" bgcolor="#003333" width="16%"><font color="#FFFFFF">�û�����:</font></td>
      <td  height="35" bgcolor="#FFFFFF" width="36%" > 
        <input type="Password" name="password" maxlength="10">
      </td>
      <td  align="right" height="35" bgcolor="#003333" width="15%"><font color="#FFFFFF">���:</font></td>
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
      <td  align="right" bgcolor="#003333" width="16%"><font color="#FFFFFF">��ϵ��ʽ:</font></td>
      <td bgcolor="#FFFFFF" width="36%" > 
        <input type="text" name="contactinfo" maxlength="50">
      </td>
      <td bgcolor="#003333" width="15%"><font color="#FFFFFF"></font></td>
      <td bgcolor="#FFFFFF" width="33%"></td>
    </tr>
    <tr> 
      <td  align="right" bgcolor="#003333" width="16%"><font color="#FFFFFF">����ʱ��:</font></td>
      <td bgcolor="#FFFFFF" width="36%" > 
        <p> 
          <select name="EndDate1">
            <% For i=2010 to 2050 
            
            %> 
            <option value="<%=i%>" <%if i=Year(Thistime) then%>selected<%end if%>><%=I%></option>
            <%next%> 
          </select>
          �� 
          <select name="EndDate2">
            <% For h=1 to 12 
            
            %> 
            <option value="<%=h%>" <%if h=Month(Thistime) then%>selected<%end if%>><%=h%></option>
            <%next%> 
          </select>
          �� 
          <select name="EndDate3">
            <% For j=1 to 31 
            
            %> 
            <option value="<%=j%>" <%if j=Day(Thistime) then%>selected<%end if%>><%=j%></option>
            <%next%> 
          </select>
          �� 
      </td>
      <td bgcolor="#003333" width="15%" ><font color="#FFFFFF"></font></td>
      <td bgcolor="#FFFFFF" width="33%" >&nbsp;</td>
    </tr>
    <tr> 
      <td  align="right" bgcolor="#003333" width="16%"><font color="#FFFFFF">Ա����:</font></td>
      <td bgcolor="#FFFFFF" width="36%" > 
       <input type="text" name="no" maxlength="16">
      </td>
      <td bgcolor="#003333" width="15%" ><font color="#FFFFFF"></font></td>
      <td bgcolor="#FFFFFF" width="33%" >&nbsp;</td>
    </tr>
  </table> 
<br>
 <input type="button" name="btnQuery" id=btnquery value="�ύ" >               
  <input type="reset" name="reset" value="����">
  <input type="button" value="����" onclick="self.history.back()">
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
 
