<!-- #include file="dbclass.asp"-->
<%
'if session("loginid")="" then
'Response.Redirect  "login.htm"
%>
<%
'dim loginid=session("loginid")
on error resume next 
dim userid
userid=Request.QueryString("userid")
dim objdml
set objdml=server.Createobject("Com_UserManage.ClsUserManage")
dim userinfo 
set userinfo=server.Createobject("adodb.recordset")
set userinfo=objdml.GetUserInfo(userid,locale,useobject)
if err.number<>0 then
	ierror=err.number
	err.clear
	set objdml=nothing
	Response.Redirect  "../../Sorry.asp?error="&ierror
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�޸ĸ����û���Ϣ</title> 
<link rel="stylesheet" href="../../style.css">
</head>


<body bgcolor=white>
<b>�����û���Ϣ</b> <br>
<form name="editlogin" Method="post" >
    
  <table border="0" width="610" cellspacing="1" cellpadding="4" bgcolor=black>
    <tr> 
      <td width="81" align="right" bgcolor=white><font color=black>ע ��  
        ID��</font></td> 
      <td width="197" bgcolor=white>  
        <input type="test" name="loginid" value="<%=userinfo("loginid")%>"> 
        <%session("loginid")=userinfo("loginid")%>
      </td> 
      <td width="82" align="right" bgcolor=white><font color=black>�Ա�</font></td> 
      <td width="209" bgcolor=white> <%if userinfo("sex")="M" then%>  
        <input type="radio" name="sex" value="M" checked> 
        ��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;  
        <input type="radio" name="sex" value="F"> 
        Ů <%else%>  
        <input type="radio" name="sex" value="M"> 
        ��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;  
        <input type="radio" name="sex" value="F" checked> 
        Ů <%end if%></td> 
    </tr> 
    <tr>  
      <td width="81" align="right" bgcolor=white><font color=black>������</font></td> 
      <td width="197" bgcolor=white>  
        <input type="test" name="name" value="<%=userinfo("name")%>"> 
      </td> 
      <td width="82" align="right" bgcolor=white><font color=black>��˾���ƣ�</font></td> 
      <td width="209" bgcolor=white>  
        <select name="CompanyID" size="1"> 
          <% 
          dim objC 
          set ObjC=server.CreateObject("Com_UserManage.ClsUserManage") 
           dim objCompany 
           set objCompany=server.CreateObject("adodb.recordset") 
           set objCompany=objC.GetCompany(Locale) 
             if err.number<>0 then 
                ierror=err.number 
                err.clear 
                set objC=nothing 
                Response.Redirect "../../Sorry.asp?Errorno="&ierror 
              end if 
              do while not objCompany.EOF 
               %>  
          <option value="<%=Objcompany("CompanyID")%>" <% 
               if objCompany("Companyid")=Userinfo("Companyid") then  
              %>selected<%End IF%>><%=ObjCompany("CompanyName")%></option> 
          <% 
             objCompany.MoveNext 
             loop 
             %>  
        </select> 
      </td> 
    </tr> 
    <tr>  
      <td width="81" align="right" height="29" bgcolor=white><font color=black>��ϵ��ʽ��</font></td> 
      <td width="197" height="29" bgcolor=white>  
        <input type="test" name="contactinfo" value="<%=userinfo("contactinfo")%>"> 
      </td> 
      <td width="82" align="right" height="29" bgcolor=white><font color=black>����(��)��</font></td> 
      <td width="209" height="29" bgcolor=white>  
        <select name="agentid" size="1"> 
          <%  
          dim objD 
          set ObjD=server.CreateObject ("Com_UserManage.ClsUserManage") 
          dim objAgent 
                 set objAgent=server.CreateObject ("adodb.recordset") 
                 set objAgent=objD.GetAgent( ) 
             if err.number<>0 then 
                ierror=err.number 
                err.clear 
                set objD=nothing 
                Response.Redirect "../../Sorry.asp?Errorno="&ierror 
              end if 
              do while not objAgent.EOF 
                 %>  
          <option value="<%=OBJagent("agentid")%>" <%if objAgent("Agentid")=Userinfo("agentid") then %>selected<%end if%>>  
          <%=ObjAgent("AgentName")%></option> 
          <%objAgent.MoveNext  
            loop 
           %>  
        </select> 
      </td> 
    </tr> 
    <tr>  
      <td width="81" align="right" bgcolor=white><font color=black>�û�����</font></td> 
      <td width="197" bgcolor=white>  
        <select name="UserObject"> 
          <%  
           dim objE 
          set ObjE=server.CreateObject ("Com_UserManage.ClsUserManage") 
           dim objUseobject 
                 set objUseobject=server.CreateObject ("adodb.recordset") 
                 set objUseObject=objE.GetUseObject( ) 
             if err.number<>0 then 
                ierror=err.number 
                err.clear 
                set objE=nothing 
                Response.Redirect "../../Sorry.asp?Errorno="&ierror 
              end if 
              objUseObject.Movefirst 
              do while not objUseObject.EOF 
            %>  
          <option value="<%=oBJuseobject("UseObject")%>" <%if objUseobject("useobject")=Userinfo("useobject") then %>selected<%end if%>>  
          <%=ObjUseobject("UseobjectName")%></option> 
          <%objUseObject.MoveNext  
            loop 
           %>  
        </select> 
      </td> 
      <td bgcolor=white width="82"><font color=black></font></td> 
      <td bgcolor=white width="209"></td> 
    </tr> 
    <tr>  
      <td width="81" bgcolor=white valign="top" > <font color=black><% 
      On Error resume next 
      dim objA 
          set ObjA=server.CreateObject ("Com_UserManage.ClsUserManage") 
      dim ObjFunction  
            set ObjFunction=server.CreateObject ("adodb.recordset") 
      set objFunction=objA.GetUserFunction (UserID,Locale) 
      if err.number<>0 then 
		ierror=err.number 
		err.clear 
		set objA=nothing 
		Response.Redirect "../../Sorry.asp?Errorno="&ierror 
      end if 
      %> </font>  
        <div ><font color=black>�û�Ȩ��:</font></div> 
      </td> 
      <td width="197" bgcolor=white valign="top"><% 
    do while not ObjFunction.EOF 
     %>  
        <div> <%=ObjFunction("FunctionName")%></div> 
        <%ObjFunction.MoveNext  
    loop 
    %> </td> 
      <td width="82" bgcolor=white valign="top" > <font color=black><% 
      dim objB 
          set ObjB=server.CreateObject ("Com_UserManage.ClsUserManage") 
      dim ObjGroup 
            set ObjGroup=server.CreateObject ("adodb.recordset") 
      set objGroup=objB.GetUserGroup (UserID,locale) 
      if err.number<>0 then 
        ierror=err.number 
        err.clear 
		set objB=nothing 
		Response.Redirect "../../Sorry.asp?Errorno="&ierror 
      end if 
      %> </font>  
        <div ><font color=black>�û�������:</font></div> 
      </td> 
      <td width="209" bgcolor=white valign="top"> <%ObjGroup.MoveFirst  
      do while not ObjGroup.EOF 
      %>  
        <div> <%=ObjGroup("GroupName")%></div> 
        <%ObjGroup.MoveNext  
     loop    
    %> </td> 
    </tr> 
    <tr>  
      <td width="81" bgcolor=white>  
        <div align="left"><font color=black>����ʱ�䣺</font></div> 
      </td> 
      <td width="197" bgcolor=white>  
        <div align="center"><%=userinfo("EndDate")%></div> 
      </td> 
      <td width="82" bgcolor=white><font color=black></font></td> 
      <td width="209" bgcolor=white>&nbsp;</td> 
    </tr> 
  </table> 
     
  <p><% 
  dim Flag 
      Flag=Request.QueryString ("Flag") 
        %> <%if flag="Del" then%>[ <a href="DELuser.asp?userid=<%=userid%>">ȷʵҪɾ����(Yes)</a> ] 
    [ <a href="userinfo.asp">����</a> ]<%end if%> <%if flag="Reset" then%>  
    [ <a href="resetuser.asp?userid=<%=userid%>">ȷʵҪ�ָ���(Yes)</a> ] [ <a href="userinfo.asp">����</a> ] 
    <%end if%> <%if flag="Pause" then%>[ <a href="pauseuser.asp?userid=<%=userid%>">ȷʵҪ��ͣ��(Yes)</a> ] 
   [ <a href="userinfo.asp">����</a> ]<%end if%> <%  
ObjFunction.Close  
objUseobject.Close 
objAgent.Close 
ObjGroup.Close 
objCompany.Close 
userinfo.Close 
set objfunction=nothing 
set objuseobject=nothing 
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
%> </p> 
</form>                        
</html> 
