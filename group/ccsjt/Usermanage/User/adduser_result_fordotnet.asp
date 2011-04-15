<!--#include file="dbclass.asp"-->
<%
   
   Set objConn = Server.CreateObject("ADODB.Connection")
   objConn.Open Application("OledbStr") 
   
   
   Set objRst=server.CreateObject ("ADODB.Recordset")
   objRst.LockType=3
   objRst.CursorType=3
   set objRst.activeConnection=objConn%>
<%
'if session("loginid")="" then
'Response.Redirect "login.htm"
%>
<%
'***********************
'输入参数:
'       UserInfo(8)
'       LoginID    (0)
'       Name       (1)
'       Sex        (2)
'       AgentID    (3)
'       CompanyID  (4)
'       ContactInfo(5) 
'       UseObject  (6)
'       password   (7)
'       EndDate    (8)
dim LoginID,Name,Sex,AgentID,CompanyID,ContactInfo,UseObject
dim Password,Enddate
    loginid=Request.Form ("loginid")
    Name=Request.Form("UserName")
    sex=Request.Form ("sex")
    Agentid=Request.Form ("Agentid")
    CompanyID=Request.Form ("Companyid")
    contactinfo=Request.Form ("Contactinfo")
    no=trim(Request.Form ("no"))
    password=Request.Form ("password")
    Enddate1=Request.Form("Enddate1")
    Enddate2=Request.Form("Enddate2")
    Enddate3=Request.Form("Enddate3")
    Enddate=enddate1&"-"&enddate2&"-"&enddate3
    objrst.Source ="select * from szairlineuser where no='"& no &"'or logid='"& loginid &"'"
    objrst.Open
    if objrst.EOF and objrst.BOF  then
    objrst.Close 
    objrst.Source ="insert into szairlineuser (name,no,logid) values ('"& name &"','"& no &"','"& loginid &"') "
    objrst.Open 
    else
    Response.Redirect "warn1.htm"
    end if
    
    
   
dim userinfo(8)
    userinfo(0)=loginid
    userinfo(1)=name
    userinfo(2)=sex
    userinfo(3)=Agentid
    userinfo(4)=companyid
    userinfo(5)=contactinfo
    userinfo(6)=Application("useobject")
    userinfo(7)=password
    userinfo(8)=enddate
    
    
    'Response.End
on error resume next
dim objdml


set objdml=CreateObject("Com_UserManage1.clsUserManage1")

userid=objdml.AddUser(userinfo)


if Err.number <>0 then
   ierror=Err.number 
   Err.Clear 
   set objdml=nothing
   response.redirect "../../Sorry.asp?Errorno="&ierror
 end if

 
'*********************
%>
<script language="javascript"> 
     alert("成功将新用户信息添加到ASP系统中！");
     window.close();
</script>
