<!--#include file="dbclass.asp"-->

<%
if trim(session("UID"))<>"" then
   dim objD1
   set ObjD1=server.CreateObject ("Com_UserManage1.clsUserManage1")
       VerifyOk=objD1.VerifyUserFunction (session("UID"),"����ϵͳ���õ�")
   if VerifyOk=false then
      session("errorNo")="000002"
      Response.Redirect "../sorry/sorry.asp"
   end if   
 else
   session("errorNo")="000001"
   Response.Redirect "../sorry/sorry.asp"
end if 
%>

<%
   
   Set objConn = Server.CreateObject("ADODB.Connection")
   objConn.Open Application("OledbStr") 
   
   
   Set objRst=server.CreateObject ("ADODB.Recordset")
   objRst.LockType=3
   objRst.CursorType=3
   set objRst.activeConnection=objConn%>
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
tempname=trim(document.edituser.name.value)
if(tempname="")then
cnt=cnt+1
cntstr=cstr(cnt)
errstr=errstr+cntstr+"."+"��������������"+chr(13)
else
document.edituser.name.value=tempname 
end if
tempcontactinfo=trim(document.edituser.contactinfo.value)
if(tempcontactinfo="")then
cnt=cnt+1
cntstr=cstr(cnt)
errstr=errstr+cntstr+"."+"������������ϵ��Ϣ"+chr(13)
else
document.edituser.contactinfo.value=tempcontactinfo
end if
enddate1=edituser.EndDate1.value
enddate2=edituser.EndDate2.value
enddate3=edituser.EndDate3.value
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
edituser.submit()
end if
end sub

-->
</script>



<%
on error resume next
dim userid
userid=Request.QueryString("userid")
dim objDml
    set ObjDml=server.CreateObject("Com_UserManage1.clsUserManage1")
dim userinfo
'dim ierror
set userinfo=server.createobject("adodb.recordset")
set userinfo=objdml.GetUserInfo(userid,Locale,useobject)
if err.number<>0 then
   ierror=err.number
   err.clear
   set ObjDml=nothing
   Response.Write "../../Sorry.asp?Errorno="&ierror
end if
%>
<%
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
      <%
      on error resume next
      dim objF
          set objF=server.CreateObject ("Com_UserManage1.clsUserManage1")
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

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�޸��û���Ϣ</title>
<link rel="stylesheet" href="../../style.css">
</head>

<body bgcolor="#FFFFFF" style="font-size:10.5pt">
<p align="left"><b><font color="#000000">�޸��û���Ϣ</font></b> 

<br><br><br>
<%
 objrst.Source="select no from szairlineuser b inner join userinfo a on a.loginid=b.logid where a.userid='" & userid & "'"
 objrst.Open
 
 
 dim pid
 if (objrst.EOF and objrst.BOF) or isnull(objrst("no")) then
     pid=""
 else
     pid=trim(objrst("no"))
 end if
 
 
%>
<form method="post" name="noquery" action="NoLinkUserid.asp">
 Ա���ţ�
 <input type="text" name="TxtNo" id="TxtNo" value=<%=pid%>>&nbsp;&nbsp;&nbsp; 
 <input type="submit"  name="BtnQueryNo" value=��ѯ>
<%objrst.Close

%> 
</form>

<form name="edituser" action="edituserout.asp?userid=<%=userid%>&loginid=<%=userinfo("loginid")%>&a=<%=Userfunction.RecordCount%>&b=<%=UserGroup.RecordCount%>" Method="post"> 
  <p align="left">&nbsp;&nbsp;&nbsp; 
  <table border="0" width="610" bgcolor="#000000" cellpadding="4" cellspacing="1">
    <tr> 
      <td width="83" align="right" bgcolor="#003333"><font color="#FFFFFF">�û�����:</font></td>
      <td width="196" bgcolor="#FFFFFF"> 
        <input type=test name="Name" value="<%=userinfo("name")%>" maxlength=50>
      </td>
      <td width="83" align="right" bgcolor="#003333"><font color="#FFFFFF">�û��Ա�:</font></td>
      <td width="207" bgcolor="#FFFFFF"> <%
      if UserInfo("sex")="M" then
      %> 
        <input type="radio" name="sex" value="M" checked>
        ��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
        <input type="radio" name="sex" value="F">
        Ů</td>
      <% else %> 
      <input type="radio" name="sex"  value="M" >
      ��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
      <input type="radio" name="sex"  value="F" checked>
      Ů <% end if%> </tr>
    <tr> 
      <td width="83" align="right" bgcolor="#003333"><font color="#FFFFFF">ע �� 
        ID:</font></td>
      <td width="196" bgcolor="#FFFFFF"> <%=userinfo("loginid")%> </td>
      <td width="83" align="right" bgcolor="#003333"><font color="#FFFFFF">��������:</font></td>
      <td width="207" bgcolor="#FFFFFF"> 
        <select name="CompanyID" size="1" >
          <%
          
          
          dim objC
          set ObjC=server.CreateObject("Com_UserManage1.clsUserManage1")
           dim objCompany
                 set objCompany=server.CreateObject("adodb.recordset")
                 
                 
                 set objCompany=objC.GetCompany(Locale)
                 
                  
             if err.number<>0 then
                ierror=err.number
                err.clear
                set objC=nothing
               
                response.redirect "../../sorry/Sorry.asp?Errorno="&ierror
             end if
              objCompany.Movefirst
              
              
              do while not objCompany.EOF
               %> 
          <option value="<%=Objcompany("CompanyID")%>" <% 
               if objCompany("Companyid")=Userinfo("Companyid") then%>selected<%End IF%>><%=ObjCompany("CompanyName")%></option>
          <%
           objCompany.MoveNext 
            loop
            
            
           %> 
        </select>
      </td>
    </tr>
    <tr> 
      <td width="83" align="right" bgcolor="#003333"><font color="#FFFFFF">��ϵ��ʽ:</font></td>
      <td width="196" bgcolor="#FFFFFF"> 
        <input type=test name="contactinfo" value="<%=userinfo("contactinfo")%>" maxlength=50>
      </td>
      <td width="83" align="right" bgcolor="#003333"><font color="#FFFFFF">����ʱ��:</font></td>
      <td width="207" bgcolor="#FFFFFF"> 
        <p> 
          <select name="EndDate1">
            <% For i=2040 to 2050 
            
            %> 
            <option value="<%=i%>" <%if i=Year(UserInfo("enddate")) then%>selected<%end if%>><%=I%></option>
            <%next%> 
          </select>
          �� 
          <select name="EndDate2">
            <% For h=1 to 12 
            
            %> 
            <option value="<%=h%>" <%if h=Month(UserInfo("enddate")) then%>selected<%end if%>><%=h%></option>
            <%next%> 
          </select>
          �� 
          <select name="EndDate3">
            <% For j=1 to 31 
            
            %> 
            <option value="<%=j%>" <%if j=Day(UserInfo("enddate")) then%>selected<%end if%>><%=j%></option>
            <%next%> 
          </select>
          ��</p>
      </td>
    </tr>
    <tr> 
      <td width="83" align="right" bgcolor="#003333"><font color="#FFFFFF">���:</font></td>
      <td width="196" bgcolor="#FFFFFF"> 
        <select name="agentid" size="1" >
          <% 
          
          
          dim objD
          set ObjD=server.CreateObject ("Com_UserManage1.clsUserManage1")
          dim objAgent
                 set objAgent=server.CreateObject ("adodb.recordset")
                 set objAgent=objD.GetAgent( )
             if err.number<>0 then
                ierror=err.number
                err.clear
                set objD=nothing
                response.redirect "../../Sorry.asp?Errorno="&ierror
              end if
              objAgent.Movefirst
              do while not objAgent.EOF
            
                 %> 
          <option value="<%=OBJagent("agentid")%>" <%if objAgent("Agentid")=Userinfo("agentid") then %>selected<%end if%>> 
          <%=ObjAgent("AgentName")%></option>
          <%objAgent.MoveNext 
            loop
           %> 
        </select>
      </td>
      <td width="83" bgcolor="#003333"></td>
      <td width="207" bgcolor="#FFFFFF"></td>
    </tr>
    <%
    objrst.Source ="select * from szairlineuser where logid='"& userinfo("loginid") &"'"
   objrst.Open 
   
       if objrst.EOF and objrst.BOF  then 
    
    %>
    <tr> 
      <td  align="right" bgcolor="#003333" width="16%"><font color="#FFFFFF">Ա����:</font></td>
      <td bgcolor="#FFFFFF" width="36%" > 
       <input type="text" name="no"   maxlength="16">
      </td>
      <td bgcolor="#003333" width="15%" ><font color="#FFFFFF"></font></td>
      <td bgcolor="#FFFFFF" width="33%" >��</td>
    </tr>
    
    <% else%>
    <tr> 
      <td  align="right" bgcolor="#003333" width="16%"><font color="#FFFFFF">Ա����:</font></td>
      <td bgcolor="#FFFFFF" width="36%" > 
       <input type="text" name="no"  value="<%=pid%>" maxlength="16">
      </td>
      <td bgcolor="#003333" width="15%" ><font color="#FFFFFF"></font></td>
      <td bgcolor="#FFFFFF" width="33%" >��</td>
    </tr>
    
    
    <%end if%>
  </table>                             
  
  <p align="left"> 
    <input type="button" name="btnQuery" id=btnQuery value=�ύ>
    <input type="button"  value=���� onclick="self.history.back()">
  </p>
  </form> 
<% 
objAgent.Close
ObjGroup.Close
objCompany.Close
userinfo.Close
UserGroup.Close 
UserFunction.Close
l=nothing 
set UserGroup=nothing
set userfunction=nothing
set objagent=nothing
set objgroup=nothing
set objcompany=nothing
set userinfo=nothing
set objdml=nothing
set ObjA=nothing
set ObjB=nothing
set oBJC=NOTHING
SET objd=nothing
set obje=nothing

%>
[ <a href="editpassword.asp?userid=<%=userid%>">�޸��û�����</a> ]
[ <a href="edituserfunction.asp?userid=<%=userid%>">���û�Ȩ��</a> ]
[ <a href="editusergroup.asp?userid=<%=userid%>">���û���</a> ]
</html>                     
