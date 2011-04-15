<!--#include file="dbclass.asp"-->

<script language="javascript"> 

window.onload=function(){ 
//function aa(){
var oldWindow;

oldWindow=window.opener;

//alert("new jump to old ");
//window.open("http://10.254.0.41/index2007.asp");
//window.close();
//alert(oldWindow.document.getElementById("txtLoginID").value);  

//try 
//{ 

document.getElementById("LoginID").value=oldWindow.document.getElementById("txtLoginID").value; 
//alert(document.getElementById("LoginID").value);

document.getElementById("Name").value=oldWindow.document.getElementById("txtName").value; 
document.getElementById("No").value=oldWindow.document.getElementById("txtEmpID").value; 
document.getElementById("Companyid").value=oldWindow.document.getElementById("ddlDept").value; 
document.getElementById("Contactinfo").value=oldWindow.document.getElementById("txtMail").value;
document.getElementById("Pwd").value=oldWindow.document.getElementById("txtPwd").value;

document.getElementById("MUID").value=oldWindow.document.getElementById("TxtMUID").value;  
document.getElementById("PUID").value=oldWindow.document.getElementById("TxtPUID").value; 

//alert(document.getElementById("Companyid").value+" "+document.getElementById("Contactinfo").value);

//alert(document.getElementById("txtNewID").value);

//alert(document.getElementById("txtUserID").value);

//document.getElementById("btnLogin").click(); 
document.forms[0].submit(); 

//} 
//catch(e) 
//{ 
//} 
} 

</script>

<%
if request.servervariables("Http_Method")="POST" then   
     
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
   
   '输入参数:LoginID
'       LoginInfo(5)
'       Name
'       Sex
'       AgentID
'       CompanyID
'       ContactInfo
'       Useobject


dim LoginInfo(4)
dim Name,Sex,AgentID,CompanyID,ContactInfo,useobject

function sqlstr(str)  
sqlstr=Replace(str,"''","~")  
end function

    Name=sqlstr(Request.Form("Name"))
    no=sqlstr(trim(Request.Form ("no")))
    sex=sqlstr(Request.Form("Sex"))
    AgentId=sqlstr(Request.Form("AgentId"))
    CompanyID=sqlstr(Request.Form("Companyid"))
    contactinfo=sqlstr(Request.Form("contactinfo"))
    Pwd=sqlstr(Request.Form("Pwd"))
    
    dim UserID
    UserID=sqlstr(Request.form("PUID"))
	dim LoginID
    loginid=sqlstr(Request.Form("loginid"))
    
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
    
    objrst.Source="update UserInfo set password='" + Pwd + "' where userid='" + userid + "'"
    objrst.Open
        
    
    logininfo(0)=name 
    logininfo(1)=sex
    logininfo(2)=AgentID
    logininfo(3)=CompanyId
    logininfo(4)=contactinfo
    
    'for i=0 to 4 
    '    Response.Write "logininfo(" + trim(i) + ")=" + logininfo(i) + "<br>"
    'next
    'Response.Write loginid + "<br>"
    'Response.Write UserID
    'Response.End
 Err.Clear 
 On Error resume next
 dim objA
     set objA=server.CreateObject("Com_UserManage.ClsUserManage")
     ierror=objA.EditLogin(LoginID,LoginInfo)
     if Err.number <>0 then
        ierror=Err.number 
        Err.Clear 
        set objA=nothing
        response.redirect "../../sorry/Sorry.asp?Errorno="&ierror
      end if
      set ObjA=nothing

dim Enddate
    Enddate=sqlstr(Request.Form ("Enddate"))
    Enddate="2050-12-31"  'cdate(enddate)
    
    dim ObjB
    set ObjB=server.CreateObject ("Com_UserManage.ClsUserManage")
    ierror=ObjB.EditUserEndDate(UserId,EndDate)
    IF Err.number <>0 then
       ierror=Err.number 
       Err.Clear 
       set objB=nothing
       response.redirect "../../sorry/Sorry.asp?Errorno="&ierror
    end if
    set objB=nothing
    
    
%>
    
    <script language="javascript"> 
       alert("成功修改ASP系统中的用户信息！");
       window.close();
    </script>
    
<%    
end if
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>修改用户信息</title>
<link rel="stylesheet" href="../../style.css">
</head>

<body bgcolor="#FFFFFF" style="font-size:10.5pt">

<br><br><br>

<form name="edituser" id="edituser"  Method="post"> 
  <p align="left">&nbsp;&nbsp;&nbsp; 
  <table  style="display:none;" border="0" width="610" bgcolor="#000000" cellpadding="4" cellspacing="1">
    <tr> 
      <td width="83" align="right" bgcolor="#003333"><font color="#FFFFFF">用户姓名:</font></td>
      <td width="196" bgcolor="#FFFFFF"> 
        <input type=test name="Name" id="Name"  maxlength=50>
      </td>
      <td width="83" align="right" bgcolor="#003333"><font color="#FFFFFF">用户性别:</font></td>
      <td width="207" bgcolor="#FFFFFF">
      <input name="sex"  id="sex" value="M"></input>
      </td>
    </tr>
    <tr> 
      <td width="83" align="right" bgcolor="#003333"><font color="#FFFFFF">注 册 ID:</font></td>
      <td width="196" bgcolor="#FFFFFF"> 
       <input type=test name="loginid" id="LoginID"> 
       </td>
      <td width="83" align="right" bgcolor="#003333"><font color="#FFFFFF">部门名称:</font></td>
      <td width="207" bgcolor="#FFFFFF"> 
        <input name="CompanyID" id="CompanyID"  type=hidden></input>
        <input name="CompanyName"  id="CompanyName"></input>
      </td>
    </tr>
    <tr> 
      <td width="83" align="right" bgcolor="#003333"><font color="#FFFFFF">联系方式:</font></td>
      <td width="196" bgcolor="#FFFFFF"> 
        <input type=test name="contactinfo" id="contactinfo"  maxlength=50>
      </td>
      <td width="83" align="right" bgcolor="#003333"><font color="#FFFFFF">结束时间:</font></td>
      <td width="207" bgcolor="#FFFFFF"> 
        <input name="EDate" id="EDate" value="2050-12-31" maxlength=50>
      </td>
    </tr>
    <tr> 
      <td width="83" align="right" bgcolor="#003333"><font color="#FFFFFF">类别:</font></td>
      <td width="196" bgcolor="#FFFFFF"> 
        <input name="agentid" value="{3E48C4DD-1FB0-40BA-BED6-856B135F3975}"  size="1" >
        </input>
      </td>
      <td width="83" bgcolor="#003333"></td>
      <td width="207" bgcolor="#FFFFFF"></td>
    </tr>
    
    <tr> 
      <td  align="right" bgcolor="#003333" width="16%"><font color="#FFFFFF">员工号:</font></td>
      <td bgcolor="#FFFFFF" width="36%" > 
       <input type="text" name="no" id="no"  maxlength="16">
      </td>
      <td bgcolor="#003333" width="15%" ><font color="#FFFFFF"></font></td>
      <td bgcolor="#FFFFFF" width="33%" >&nbsp;</td>
    </tr>
    
    <tr> 
      <td  align="right" bgcolor="#003333" width="16%"><font color="#FFFFFF">员工号:</font></td>
      <td bgcolor="#FFFFFF" width="36%" > 
       <input type="text" name="MUID" id="MUID"  maxlength="16">
       <input type="text" name="PUID" id="PUID"  maxlength="16">
       <input type="text" name="Pwd" id="Pwd"  maxlength="16">
      </td>
      <td bgcolor="#003333" width="15%" ><font color="#FFFFFF"></font></td>
      <td bgcolor="#FFFFFF" width="33%" >&nbsp;</td>
    </tr>
     
  </table>                             
  
 </form> 

</html>                     

