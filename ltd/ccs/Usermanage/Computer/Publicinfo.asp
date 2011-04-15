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
set userinfo=objdml.GetComputerInfo(userid,locale,useobject)
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
<title>修改个人用户信息</title> 
<link rel="stylesheet" href="../../style.css">
</head>


<body bgcolor="#FFFFFF">
<b>个人用户信息</b> <br>
<form name="editlogin" Method="post" >
    
  <table border="0" width="610" cellspacing="1" cellpadding="4" bgcolor="#000000">
    <tr> 
      <td width="81" align="right" bgcolor="#003333"><font color="#FFFFFF">注 册  
        ID：</font></td> 
      <td width="197" bgcolor="#FFFFFF">  
        <input type="test" name="loginid" value="<%=userinfo("loginid")%>"> 
      </td> 
     <td width="82" align="right" bgcolor="#003333"><font color="#FFFFFF">公司名称：</font></td> 
      <td width="209" bgcolor="#FFFFFF">  
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
      <td width="81" align="right" height="29" bgcolor="#003333"><font color="#FFFFFF">联系方式：</font></td> 
      <td width="197" height="29" bgcolor="#FFFFFF">  
        <input type="test" name="contactinfo" value="<%=userinfo("contactinfo")%>"> 
      </td> 
      <td width="82" align="right" height="29" bgcolor="#003333"><font color="#FFFFFF">代理(商)：</font></td> 
      <td width="209" height="29" bgcolor="#FFFFFF">  
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
      <td width="81" align="right" bgcolor="#003333"><font color="#FFFFFF">用户对象：</font></td> 
      <td width="197" bgcolor="#FFFFFF">  
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
      <td bgcolor="#003333" width="82"><font color="#FFFFFF"></font></td> 
      <td bgcolor="#FFFFFF" width="209"></td> 
    </tr> 
    <tr>  
      <td width="81" bgcolor="#003333" valign="top" > <font color="#FFFFFF"><% 
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
        <div ><font color="#FFFFFF">用户权限:</font></div> 
      </td> 
      <td width="197" bgcolor="#FFFFFF" valign="top"><% 
    do while not ObjFunction.EOF 
     %>  
        <div> <%=ObjFunction("FunctionName")%></div> 
        <%ObjFunction.MoveNext  
    loop 
    %> </td> 
      <td width="82" bgcolor="#003333" valign="top" > <font color="#FFFFFF"><% 
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
        <div ><font color="#FFFFFF">用户所在组:</font></div> 
      </td> 
      <td width="209" bgcolor="#FFFFFF" valign="top"> <%ObjGroup.MoveFirst  
      do while not ObjGroup.EOF 
      %>  
        <div> <%=ObjGroup("GroupName")%></div> 
        <%ObjGroup.MoveNext  
     loop    
    %> </td> 
    </tr> 
    <tr>  
      <td width="81" bgcolor="#003333">  
        <div align="left"><font color="#FFFFFF">结束时间：</font></div> 
      </td> 
      <td width="197" bgcolor="#FFFFFF">  
        <div align="center"><%=userinfo("EndDate")%></div> 
      </td> 
      <td width="82" bgcolor="#003333"><font color="#FFFFFF"></font></td> 
      <td width="209" bgcolor="#FFFFFF">&nbsp;</td> 
    </tr> 
  </table> 
     
  <p><% 
  dim Flag 
      Flag=Request.QueryString ("Flag") 
        %> <%if flag="Del" then%>[ <a href="DELuser.asp?userid=<%=userid%>">确实要删除吗？(Yes)</a> ] 
    [ <a href="userinfo.asp">返回</a> ]<%end if%> <%if flag="Reset" then%>  
    [ <a href="resetuser.asp?userid=<%=userid%>">确实要恢复吗？(Yes)</a> ] [ <a href="userinfo.asp">返回</a> ] 
    <%end if%> <%if flag="Pause" then%>[ <a href="pauseuser.asp?userid=<%=userid%>">确实要暂停吗？(Yes)</a> ] 
   [ <a href="userinfo.asp">返回</a> ]<%end if%> <%  
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
