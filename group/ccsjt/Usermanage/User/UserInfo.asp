<!--#include file="dbclass.asp"-->
<html>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Sub btnQuery_onclick
	frminfo.action = "Userinfo.asp?tiaojian=" & frminfo.tiaojian.value & "&textfield=" & frminfo.textfield.value
	frminfo.submit  
End Sub
-->
</SCRIPT>

<head>
 <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function submit1_onclick() {
temp.text1.value="History"
temp.submit()
}

function submit2_onclick() {
temp.text1.value="Edit"
temp.submit()
}

function submit3_onclick() {
temp.text1.value="Del"
temp.submit ()
}

//-->
</SCRIPT>
<link rel="stylesheet" href="../../style.css">
 <title></title>
</head>
 <body>
<table width="90%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td><b><font color="#0000FF" size="5">�û���Ϣ��¼</font></b>
        <form name="frminfo" method="post" action="userinfo.asp" >
          <div align="left">
          <table width="610" border="0" cellspacing="2" cellpadding="4" height="15" bgcolor="#006699">
            <tr>
              <td><font color="#FFFFFF">��ѯ������</font> 
                <select name="tiaojian">
                  <option value="Name">����</option>
                  <option value="LoginID">��¼��</option>
                  <option value="CompanyName">��������</option>
                  <option value="AgentName">���</option>
                </select>
                <input type="text" name="textfield">
                <input type="button" name="BtnQuery" id=BtnQuery  value="��ʼ��ѯ">
                <input type="reset" name="reset" value="����">
              </td>
            </tr>
          </table>
          
        </div>
        </form>
      </td>
    </tr>
    <tr>
      
    <td><%
  On Error resume next 
  dim useobject
  useobject="AMS" 'application("useobject")
  const DbClassA=7

dim tiaojian,textfield
    tiaojian=Request.QueryString ("tiaojian")
    textfield=Request.QueryString ("textfield")
    textfield=trim(textfield)

dim ObjDml
    set ObjDml=server.CreateObject ("Com_UserManage1.ClsUserManage1")

dim ObjUserInfo 
    set ObjUserInfo=server.CreateObject ("adodb.recordset")
    set ObjUserInfo=ObjDml.SearchUserInfo (Useobject,textfield,tiaojian)
if err.number<>0 then
   Ierror=err.number
   err.clear
   set objdml=nothing
   'Response.End 
   response.redirect "../../Sorry.asp?Errorno=" &ierror
  
end if  
 
const PAGE_SIZE = 8
objUserInfo.PageSize = PAGE_SIZE
Dim iCurrentPage

if CInt(Request.QueryString("PageNo"))>=1 and CInt(Request.QueryString("PageNo"))<=objUserInfo.PageCount then
	iCurrentPage = CInt(Request.QueryString("PageNo"))
else
	iCurrentPage =1
end if

If not objUserInfo.EOF Then
	objUserInfo.AbsolutePage = iCurrentPage
	If iCurrentPage > 1 Then
		Response.Write  "<A href='Userinfo.asp?PageNo=" & (iCurrentPage-1) & "& Userid=" & Request.QueryString("Userid") &  "'>��һҳ</a>&nbsp;&nbsp;"
	    Response.Write "<a href='Userinfo.asp?PageNo=" & 1  & "& Userid=" & Request.QueryString("Userid") &  "'>��ҳ</a>&nbsp;&nbsp;"

	End If
	If iCurrentPage < objUserInfo.PageCount Then
		Response.Write "<A href='Userinfo.asp?PageNo=" & (iCurrentPage+1) & "&Userid=" & Request.QueryString("Userid") &  "'>��һҳ</a>&nbsp;&nbsp;"
	    Response.Write "<a href='Userinfo.asp?PageNo=" & ObjUserInfo.PageCount  & "& Userid=" & Request.QueryString("Userid") &  "'>βҳ</a>&nbsp;&nbsp;"
	End If
%> �� <%=iCurrentPage%> / <%=objUserInfo.PageCount%> ҳ<br>
      <form name="temp" method="post" action="sele.asp" > 
        <table width="610" border="0" bgcolor="#003333" cellspacing="1" cellpadding="4"> 
          <tr width=100% bgcolor="#003333">  
            <td height="2" width="29" ><font color="#FFFFFF"> ѡ��</font></td> 
            <td height="2" width="54"><font color="#FFFFFF">��¼��</font></td> 
            <td height="2" width="50"><font color="#FFFFFF">����</font></td> 
            <td height="2" width="31"><font color="#FFFFFF">�Ա�</font></td> 
            <td height="2" width="79"><font color="#FFFFFF">���</font></td> 
            <td height="2" width="79"><font color="#FFFFFF">��������</font></td> 
            <td height="2" width="69"><font color="#FFFFFF">��ϵ��Ϣ</font></td> 
            <td height="2" width="62"><font color="#FFFFFF">����ʱ��</font></td> 
            <td height="2" width="63"><font color="#FFFFFF">��ͣ���ָ�</font></td> 
          </tr> 
          <% 
	dim i 
	For i=1 to PAGE_SIZE 
%>

<% 
if ObjUserinfo("status")="E" then 
%>  
          <tr bgcolor="#FFFFFF">  
            <td width="29">  
              <input type=radio name=user value="<%=objuserinfo("userid")%>" <%if i=1 then%> checked <%end if%> > 
            </td> 
            <td width="54"><%=ObjUserInfo("LoginID")%></td> 
            <td width="50"><%=ObjUserInfo("Name")%></td> 
            <td width="31"><%=ObjUserInfo("Sex")%></td> 
            <td width="79"><%=ObjUserInfo("AgentName")%></td> 
            <td width="79"><%=ObjUserInfo("CompanyName")%></td> 
            <td width="69"><%=ObjUserInfo("ContactInfo")%></td> 
            <td width="62"><%=ObjUserInfo("EndDate")%></td> 
            <td width="63"><a href="Publicinfo.asp?UserID=<%=ObjUserInfo("UserID")%>&Flag=Pause">��ͣ</a></td> 
          </tr> 
          <%End if %> <%if ObjUserinfo("status")="S" then%>  
          <tr bgcolor="#FFFFFF">  
            <td width="29">  
              <input type=radio name=user value=<%=objuserinfo("userid")%> <%=objuserinfo("userid")%> <%if i=1 then%> checked <%end if%>> 
            </td> 
            <td width="54"><%=ObjUserInfo("LoginID")%></td> 
            <td width="50"><%=ObjUserInfo("Name")%></td> 
            <td width="31"><%=ObjUserInfo("Sex")%></td> 
            <td width="79"><%=ObjUserInfo("AgentName")%></td> 
            <td width="79"><%=ObjUserInfo("CompanyName")%></td> 
            <td width="69"><%=ObjUserInfo("ContactInfo")%></td> 
            <td width="62"><%=ObjUserInfo("EndDate")%></td> 
            <td width="63"><a href="Publicinfo.asp?UserID=<%=ObjUserInfo("UserID")%>&Flag=Reset">�ָ�</a></td> 
          </tr> 
          <%end if 
		objUserInfo.movenext 
		If objUserInfo.EOF Then 
			Exit For 
		End If 
	next%>  
        </table> 
        <a href="Adduser.asp"> </a>  
        <table width="610" border="0" cellspacing="0" cellpadding="6"> 
          <tr> 
            <td><% 
Else 
	Response.Write "��ǰû�м�¼" 
End If 
set objDML=nothing 
%>               <input type="hidden" id=text1 name=text1> 
              <input type="submit" value="��ʷ" id=submit1 name=submit1 language=javascript onClick="return submit1_onclick()"> 
              <input type="submit" value="�޸�" id=submit2 name=submit2 language=javascript onClick="return submit2_onclick()"> 
              <input type="submit" value="ɾ��" id=submit3 name=submit3 language=javascript onClick="return submit3_onclick()"> 
            </td> 
            <td align="center">[ <a href="Adduser.asp">����û���¼</a> ] [ <a href="browselogininfo.asp">�鿴ɾ���û���¼</a>  
              ] </td> 
          </tr> 
        </table> 
      </form> 
</td> 
    </tr> 
  </table> 
  </body> 
</html>