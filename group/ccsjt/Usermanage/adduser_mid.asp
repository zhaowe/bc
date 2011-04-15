<html>
	<head>
		<title>Finance Excutive Information System</title>
		<script language="javascript"> 

window.onload=function(){ 
//function aa(){
var oldWindow; 
if(document.getElementById("txtUserID").value=="") 
{ 

//try 
//{ 
oldWindow=window.opener; 

//alert(oldWindow.document.getElementById("ToASP1_UID").value); 

document.getElementById("txtUserID").value=oldWindow.document.getElementById("txtLoginID").value; 
document.getElementById("txtPwd").value=oldWindow.document.getElementById("txtPwd").value; 
document.getElementById("txtUName").value=oldWindow.document.getElementById("txtName").value; 

alert(document.getElementById("txtUserID").value);

//document.getElementById("btnLogin").click(); 
document.forms[0].submit(); 

//} 
//catch(e) 
//{ 
//} 
} 
} 
		</script>

		<% 

if request.servervariables("Http_Method")="POST" then   

function sqlstr(str)  
sqlstr=Replace(str,"''","~")  
end function 
 
loginID=sqlstr(request("userID"))  
password=sqlstr(request("Pwd")) 
name=sqlstr(request("username"))
sex="F"
Agentid="{5E28769D-8FDA-46E9-967D-437A5379AFC5}"
companyid="25"
contactinfo=""
useobject="contan"
enddate="2050-6-30" 

%>




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
'ÊäÈë²ÎÊý:
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
<%
'Response.Write userid
Response.Redirect "userinfo.asp"
%>



<% end if  %>


	</head>
	<body>
		<form METHOD="post" NAME="login" ID="Form1">
			<table style="display:none;">
				<tr>
					<td><INPUT id="txtUserID" name="userID" MAXLENGTH="20"></td>
				</tr>
				<tr>
					<td><input NAME="Pwd" MAXLENGTH="20" id="txtPwd"></td>
				</tr>
				<tr>
					<td><input NAME="username" id="txtUName"></td>
					</td></tr>
			</table>
		</form>
	</body>
</html>
