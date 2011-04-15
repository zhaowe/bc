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
document.getElementById("txtEmpID").value=oldWindow.document.getElementById("txtEmpID").value; 

document.getElementById("txtCompanyid").value=oldWindow.document.getElementById("ddlDept").value; 
document.getElementById("txtContactinfo").value=oldWindow.document.getElementById("txtMail").value; 

document.getElementById("txtNewID").value=oldWindow.document.getElementById("txtNewID").value; 

//alert(document.getElementById("txtCompanyid").value+" "+document.getElementById("txtContactinfo").value);

//txtEmpID txtNewID

//alert(document.getElementById("txtNewID").value);

//alert(document.getElementById("txtUserID").value);

//document.getElementById("btnLogin").click(); 
document.forms[0].submit(); 

//} 
//catch(e) 
//{ 
//} 
} 
} 
</script>

		
		
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


if request.servervariables("Http_Method")="POST" then   

function sqlstr(str)  
sqlstr=Replace(str,"''","~")  
end function 
 
loginID=sqlstr(request("userID"))  
password=sqlstr(request("Pwd")) 
name=sqlstr(request("username"))
no=sqlstr(request("txtEmpID"))
companyid=sqlstr(request("txtCompanyid"))
contactinfo=sqlstr(request("txtContactinfo"))


sex="M"
Agentid="{3E48C4DD-1FB0-40BA-BED6-856B135F3975}"
Application("useobject")="AMS"
enddate="2050-12-31" 

mark=sqlstr(request("txtNewID"))

'loginID="lizz6234"
'password="abcd1234"
'name="lipenglipeng"
'no="510000"
'sex="M"
'Agentid="{3E48C4DD-1FB0-40BA-BED6-856B135F3975}"
'companyid=5
'contactinfo="6275"
'Application("useobject")="AMS"
'enddate="2050-6-30" 
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
Err.Clear

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
    
    'for i=0 to 8
    '    Response.Write  "(" + trim(i) + ")" + userinfo(i)+ "<br>"
     
    'next  
    'Response.End
     
    objrst.Source ="select * from szairlineuser where no='"& no &"'or logid='"& loginid &"'"
    'Response.Write objrst.Source
    'Response.End
    
    
    objrst.Open
    
    if objrst.EOF and objrst.BOF  then
       objrst.Close 
       objrst.Source ="insert into szairlineuser (name,no,logid) values ('"& name &"','"& no &"','"& loginid &"') "
       objrst.Open 
    else
       Response.Redirect "warn1.htm"
    end if
   

on error resume next
dim objdml


'for i=0 to 8 
'   Response.Write "userinfo(" + trim(i) + ")=" + userinfo(i) + "<br>"
'next
'Response.End

set objdml=CreateObject("Com_UserManage1.clsUserManage1")
userid=objdml.AddUser(userinfo)

if Err.number <>0 then
   Response.Write err.Description + "<br>"
   Response.Write err.Source
   Response.End
   ierror=Err.number 
   Err.Clear 
   set objdml=nothing
   response.redirect "../../sorry/Sorry.asp?Errorno="&ierror
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
		<form METHOD="post" NAME="login" ID="Form1" >
			<table style="display:none;">
				<tr>
					<td><INPUT id="txtUserID" name="userID" MAXLENGTH="20"></td>
				</tr>
				<tr>
					<td><input NAME="Pwd" MAXLENGTH="20" id="txtPwd"></td>
				</tr>
				<tr>
					<td><input NAME="username" id="txtUName"></td>
				</td>
				</tr>
				<tr>
					<td><input NAME="txtEmpID" id="txtEmpID"></td>
				</tr>
			
				<tr>
					<td><input NAME="txtCompanyid" id="txtCompanyid"></td>
				</tr>
				<tr>
					<td><input NAME="txtContactinfo" id="txtContactinfo"></td>
				</tr>
				<tr>
					<td><input NAME="txtNewID" id="txtNewID"></td>
				</tr>
			</table>
		</form>
		
		
	</body>
</html>
