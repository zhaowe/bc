<%@ Language=VBScript %>
<!--#include file="../../include/UserCheck.asp"-->

<%
DateStr=Date()
Year_Str=year(DateStr)
Month_Str=month(DateStr)
Day_Str=day(DateStr)
Today=Year_Str & "/" & Month_Str & "/" & Day_Str

set rsAgentType=server.CreateObject("ADODB.recordset")
set objCredit=server.CreateObject("com_Agent.clsAgent")
set objDML=server.CreateObject("com_DML.clsDML")

'get recordset from database
set rsAgentType=objCredit.GetAgentTypeList(application("Locale"))
if Err.number<>0 then  '���˵����
   ErrNo=Err.number 
   Err.Clear 
   set objCredit=nothing
   set rsAgentType=nothing
   Response.Redirect("../../sorry.asp?ErrorNo=" + cstr(ErrNo))
end if


Set rsAgentList = objCredit.GetAllAgentInfo("V",Application("UseObject"),Application("Locale"))
if Err.number<>0 then  '�ϼ��ڵ㼯��
   ErrNo=Err.number 
   Err.Clear 
   set objCredit=nothing
   set rsAgentList=nothing
   Response.Redirect("../../sorry.asp?ErrorNo=" + cstr(ErrNo))
end if



%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT LANGUAGE=javascript>
<!--
function submit1_Click() {

if(document.form1.text1.value == ""){
			alert("���ڴ����Office��������������!")
			return false
		}	
if(document.form1.text2.value == ""){
			alert("���ڴ����������������ô����ȫ��!")
			return false
		}	
if(document.form1.text3.value == ""){
			alert("���ڴ��������������ô������!")
			return false
		}	
		
if(document.form1.text5.value == ""){
			alert("������ϵ��ַ��������ô���㵥λ��ϵ��ַ!")
			return false
		}	
if(document.form1.text6.value == ""){
			alert("������ϵ����������ô���㵥λ��ϵ��!")
			return false
		}	
if(document.form1.text7.value == ""){
			alert("������ϵ�绰��������ô���㵥λ��ϵ�绰!")
			return false
		}	
if(document.form1.text8.value == ""){
			alert("�������ڳ�����������ô�������ڳ���!")
			return false
		}	
if(document.form1.text13.value == ""){
			alert("���ڲ����������������ô�����뻧�Ĳ���Ա!")
			return false
		}	
		
		document.form1.submit()
		return true}
//-->
</SCRIPT>

<link rel="stylesheet" href="../../style.css">
</HEAD>
<BODY bgcolor="#FFFFFF">
<FORM action="NewAgentOk.asp" method=post id=form1 name=form1>
  <p><b><font color="#0000FF">���������(��������)</font></b></p>
  <TABLE border=0 cellPadding=4 cellSpacing=1 width="610" bgcolor="#000000">
    <TR> 
      <TD style="WIDTH: 25%" width="25%" bgcolor="#003333"><font color="#FFFFFF">�����Office:</font></TD>
      <TD style="WIDTH: 25%" width="25%" bgcolor="#FFFFFF"> 
        <INPUT id=text1 
      name=AgentOffice style="HEIGHT: 22px; WIDTH: 88px" >
      </TD>
      <TD style="WIDTH: 25%" width="25%" bgcolor="#003333"><font color="#FFFFFF">��������</font></TD>
      <TD bgcolor="#FFFFFF"> 
        <SELECT id=select1 name=AgentType size=1 style="HEIGHT: 29px; WIDTH: 76px">
          <% for i=1 to rsAgentType.RecordCount %> 
          <OPTION <%if ucase(rsAgentType("AgentType"))="J"  then%> selected <%end if%> value=<%=rsAgentType("AgentType")%>><%=rsAgentType("AgentTypeName")%></OPTION>
          <% rsAgentType.MoveNext 
          Next%> 
        </SELECT>
      </TD>
    </TR>
    <TR> 
      <TD bgcolor="#003333"><font color="#FFFFFF">���������</font></TD>
      <TD bgcolor="#FFFFFF"> 
        <INPUT id=text2 name=AgentName
      style="HEIGHT: 22px; WIDTH: 88px" >
      </TD>
      <TD bgcolor="#003333"><font color="#FFFFFF">�������</font></TD>
      <TD bgcolor="#FFFFFF"> 
        <INPUT id=text3 name=AgentShortName
      style="HEIGHT: 22px; WIDTH: 88px" >
      </TD>
    </TR>
    <TR> 
      <TD bgcolor="#003333"><font color="#FFFFFF">�ϼ������</font></TD>
      <TD bgcolor="#FFFFFF"> 
        <SELECT id=select1 name=FAgentID size=1  style="HEIGHT: 29px; WIDTH: 76px">
          <% for i=1 to rsAgentList.RecordCount %> 
          <Option <%if rsAgentList("AgentType")="H" then%> selected <%end if%>  value="<%=rsAgentList("AgentID")%>"> 
          <%=rsAgentList("AgentName")%></OPTION>
          <%rsAgentList.MoveNext 
    next%> 
        </SELECT>
      </TD>
      <TD bgcolor="#003333"><font color="#FFFFFF">��ϵ��ַ</font></TD>
      <TD bgcolor="#FFFFFF"> 
        <INPUT id=text5 name=lxrAdd 
      style="HEIGHT: 22px; WIDTH: 88px" >
      </TD>
    </TR>
    <TR> 
      <TD bgcolor="#003333"><font color="#FFFFFF">��ϵ��</font></TD>
      <TD bgcolor="#FFFFFF"> 
        <INPUT id=text6 name=lxrName 
      style="HEIGHT: 22px; WIDTH: 88px" >
      </TD>
      <TD bgcolor="#003333"><font color="#FFFFFF">��ϵ�绰</font></TD>
      <TD bgcolor="#FFFFFF"> 
        <INPUT id=text7 name=lxrPhone
      style="HEIGHT: 22px; WIDTH: 88px" >
      </TD>
    </TR>
    <TR> 
      <TD bgcolor="#003333"><font color="#FFFFFF">���ڳ���</font></TD>
      <TD bgcolor="#FFFFFF"> 
        <INPUT id=text8 name=AgentCity
      style="HEIGHT: 22px; WIDTH: 88px" >
      </TD>
      <TD bgcolor="#003333"><font color="#FFFFFF">��������</font></TD>
      <TD bgcolor="#FFFFFF"> 
        <INPUT id=text9 name=OpenBank 
      style="HEIGHT: 22px; WIDTH: 88px" >
      </TD>
    </TR>
    <TR> 
      <TD bgcolor="#003333"><font color="#FFFFFF">�����ʺ�</font></TD>
      <TD bgcolor="#FFFFFF"> 
        <INPUT id=text10 name=OpenAccount 
      style="HEIGHT: 22px; WIDTH: 88px" >
      </TD>
      <TD bgcolor="#003333"><font color="#FFFFFF">Э�����</font></TD>
      <TD bgcolor="#FFFFFF"> 
        <INPUT id=text11 name=ProtocalNo
      style="HEIGHT: 22px; WIDTH: 88px" >
      </TD>
    </TR>
    <TR> 
      <TD bgcolor="#003333"><font color="#FFFFFF">Э������</font></TD>
      <TD bgcolor="#FFFFFF"> 
        <INPUT id=text12 name=ProtocalDate
      style="HEIGHT: 22px; WIDTH: 88px" value=<%=Today%>>
      </TD>
      <TD bgcolor="#003333"><font color="#FFFFFF">��ʵѡ��</font></TD>
      <TD bgcolor="#FFFFFF"> 
        <SELECT id=select2 name=AgentEntity 
      style="HEIGHT: 22px; WIDTH: 77px">
          <OPTION   selected  value="T">ʵ��</OPTION>
          <OPTION   value="F">���</OPTION>
        </SELECT>
      </TD>
    </TR>
  </TABLE>
  <p>
    <input type="submit" value="���������" id=submit1 name=submit1 onClick="return submit1_Click()">
     <input type="button" value="ȡ��" id=submit1 name=submit12 onClick="self.history.back()">
  </p>
</FORM>

</BODY>
</HTML>
