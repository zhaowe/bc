<!--#include file="dbclass.asp"-->
<%
   
   Set objConn = Server.CreateObject("ADODB.Connection")
   objConn.Open Application("OledbStr") 
   
   
   Set objRst=server.CreateObject ("ADODB.Recordset")
   objRst.LockType=3
   objRst.CursorType=3
   set objRst.activeConnection=objConn
   
    Set objRst1=server.CreateObject ("ADODB.Recordset")
   objRst1.LockType=3
   objRst1.CursorType=3
   set objRst1.activeConnection=objConn
   %>
<%'ÊäÈë²ÎÊý:LoginID
'       LoginInfo(5)
'       Name
'       Sex
'       AgentID
'       CompanyID
'       ContactInfo
'       Useobject
%>
<%'Edit table of LoginInfo
dim UserID
    UserID=Request.QueryString ("UserID")
dim LoginID
    loginid=Request("loginid")
dim LoginInfo(4)
dim Name,Sex,AgentID,CompanyID,ContactInfo,useobject
    Name=Request.Form("Name")
    no=trim(Request.Form ("no"))
    sex=Request.Form("Sex")
    AgentId=Request.Form("AgentId")
    CompanyID=Request.Form("Companyid")
    contactinfo=Request.Form("contactinfo")
    objrst.Source ="select * from szairlineuser where logid='"& loginid &"'"
    objrst.Open 
    if objrst.BOF and objrst.EOF  then
    objrst.Close 
    objrst.Source ="select * from szairlineuser where no='"& no &"'"
    objrst.Open
    if objrst.BOF and objrst.EOF  then
    objrst1.Source ="insert into szairlineuser (name,no,logid) values ('"& name &"','"& no &"','"& loginid &"') "
    objrst1.Open 
    objrst.Close 
    else
     Response.Redirect "warn2.htm"
    end if 
    else
    objrst.Close 
    objrst.Source ="select * from szairlineuser where no='"& no &"' and logid<>'"& loginid &"'"
    objrst.Open
    if objrst.BOF and objrst.EOF  then
   objrst1.Source ="update szairlineuser set name='"& name &"',no='"& no &"' where logid='"& loginid &"'"
    objrst1.Open 
    objrst.Close 
    else
     Response.Redirect "warn2.htm"
    end if 
   
    end if
    
    
    logininfo(0)=name 
    logininfo(1)=sex
    logininfo(2)=AgentID
    logininfo(3)=CompanyId
    logininfo(4)=contactinfo

 On Error resume next
 dim objA
     set objA=server.CreateObject("Com_UserManage.ClsUserManage")
     ierror=objA.EditLogin(LoginID,LoginInfo)
     if Err.number <>0 then
        ierror=Err.number 
        Err.Clear 
        set objA=nothing
        response.redirect "../../Sorry.asp?Errorno="&ierror
      end if
      set ObjA=nothing
     %>
<%'Edit table of UserInfo
dim Enddate
dim Enddate1,Enddate2,Enddate3
    Enddate1=Request.Form ("Enddate1")
    Enddate2=Request.Form ("Enddate2")
    Enddate3=Request.Form ("Enddate3")
    Enddate=Enddate1 & "-" & enddate2 & "-" & enddate3
    Enddate=cdate(enddate)
    
dim ObjB
    set ObjB=server.CreateObject ("Com_UserManage.ClsUserManage")
   ierror=ObjB.EditUserEndDate(UserId,EndDate)
    IF Err.number <>0 then
       ierror=Err.number 
       Err.Clear 
       set objB=nothing
       response.redirect "../../Sorry.asp?Errorno="&ierror
    end if
    set objB=nothing
%> 

<%Response.Redirect "EditUser.asp?UserId="&UserId
'******************************************************
%>


