<!--#include file="../../include/UserCheck.asp"-->
<%
i_Page=request("Page")

on error resume next
set rsAgentlist=server.CreateObject("ADODB.recordset")
set objCredit=server.CreateObject("com_Agent.clsAgent")
objCredit.Locale=application("Locale")
objCredit.UseObject=application("UseObject")
set rsAgentlist=objCredit.GetAgentInfo("ALL","ALL","ALL")

 If Err.Number<>0 then
   ErrNo=Err.Number
   Err.clear
   set objCredit=nothing
   Response.Redirect("../../sorry.asp?ErrorNo=" + cstr(ErrNo))
 End if

rsAgentlist.PageSize=8
if i_Page<>"" then
  if i_Page<1 then
    i_Page=1
  else
   if cint(i_Page)>cint(rsAgentlist.PageCount)  then
     i_Page=rsAgentlist.PageCount 
   end if
  end if
else
 i_Page=1
end if
 
if not rsAgentlist.EOF then
 rsAgentlist.AbsolutePage = i_Page
end if


%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--

function modiAgent_onclick() {
	document.Agentform.action="UpdateAgent.asp"
	document.Agentform.submit()
	return true
}

function cancelAgent_onclick() {
	document.Agentform.action="viewAgentInfo.asp"  
	document.Agentform.text1.value="True"
	document.Agentform.submit()
	return true
}

function setcredit_onclick() {
	document.Agentform.action="../Credit/SetCredit.asp"
	document.Agentform.submit()
	return true
}
function viewcredit_onclick() {
	document.Agentform.action="../Credit/ViewCreditInfo.asp"
	document.Agentform.submit()
	return true
}
//-->
</script>
<link rel="stylesheet" href="../../style.css">
</head>
<body bgcolor="#FFFFFF">
<p> 
<p><b><font color="#0000FF" size="5">�ɹ�ѡ��Ĵ������Ϣ�б�</font></b></p>
<%if i_Page<>1 then%>
<p><a href="viewAgentList.asp?Page=1">��ҳ</a>&nbsp;&nbsp;<a href="viewAgentList.asp?Page=<%=i_Page-1%>">��ҳ</a> 
  <%end if%> <%if i_Page-rsAgentlist.PageCount <>0 then%> <a href="viewAgentList.asp?Page=<%=i_Page+1%>">��ҳ</a>&nbsp;&nbsp;<a href="viewAgentList.asp?Page=<%=rsAgentlist.PageCount %>">βҳ</a> 
  <%end if%> ��<%=i_Page%>ҳ ��<%=rsAgentlist.PageCount %>ҳ <%if not rsAgentlist.EOF <>0 then%> 
</p>
<form method="post" action="viewAgentInfo.asp" name="Agentform" id=Agentform>
 
  <table border="0" width="610" height="20" cellspacing="1" cellpadding="4" bgcolor="#000000">
    <tr bgcolor="#003333"> 
      <td width="30" ><font color="#FFFFFF">ѡ��</font></td>
      <td width="60" height="1" ><font color="#FFFFFF">Office���</font></td>
      <td width="80" height="1" ><font color="#FFFFFF">���������</font></td>
      <td width="30" height="1" ><font color="#FFFFFF">���</font></td>
      <td width="60" height="1" ><font color="#FFFFFF">����</font></td>
      <td width="50" height="1" ><font color="#FFFFFF">Tel</font></td>
      <td width="50" height="1" ><font color="#FFFFFF">�ϼ���</font></td>
      <td width="60" height="1" ><font color="#FFFFFF">�����ʻ�</font></td>
      <td width="63" height="1" ><font color="#FFFFFF">��ͣ���ָ�</font></td>
    </tr>
    <%For i=1 to rsAgentList.PageSize  %> 
    <tr bgcolor="#FFFFFF"> 
      <td width="17"> 
        <input type="radio" id="radio1" <%if i=1 then%> checked <%end if%> value="<%=rsAgentList("AgentID")%>" name="Agent_radio" >
      </td>
      <td width="60" height="1" ><%=rsAgentList("AgentOffice")%></td>
      <td width="80" height="1" ><%=rsAgentList("AgentName")%></td>
      <td width="80" height="1" ><%=rsAgentList("AgentTypeName")%></td>
      <td width="60" height="1" ><%if len(trim(rsAgentList("AgentCity")))=0 then%>-<%else%><%=rsAgentList("AgentCity")%><%end if%></td>
      <td width="50" height="1" ><%=rsAgentList("lxrPhone")%></td>
      <td width="50" height="1" ><%if isnull(rsAgentList("FAgentName")) then%>-<%else%><%=rsAgentList("FAgentName")%><%end if%></td>
      <td width="60" height="1" ><%=rsAgentList("CreditID")%></td>
      <td width="63" height="1" > <%if rsAgentList("DealStatus")="V" then%> <a href="viewAgentInfo.asp?Agent_radio=<%=rsAgentList("AgentID")%>&Flag=Pause">��ͣ</a> 
        <%else%> <%if rsAgentList("DealStatus")="P" then%> <a href="viewAgentInfo.asp?Agent_radio=<%=rsAgentList("AgentID")%>&Flag=resume">�ָ�</a> 
        <%end if%> <%end if%> </td>
    </tr>
    <%rsAgentList.MoveNext 
  If rsAgentList.EOF Then
	Exit For
  End If
	next%> 
  </table>
<%else%>
<p>��ǰû�м�¼</p>
<%end if%>
<br>
  <table width="610" border="0" cellspacing="1" cellpadding="4">
    <tr>
      <td width="515"> 
        <input type="hidden" id=text1 name=DelFlag>
        <input type=submit name=modiAgent id=modiAgent value="�޸Ĵ������Ϣ"  language="javascript" onClick="return modiAgent_onclick()">
        <input type=submit name=cancelAgent id=cancelAgent value="�رմ����"  language="javascript" onClick="return cancelAgent_onclick()">
        <input type=submit name=setcredit id=setcredit value="ָ�����������ʻ�"  language="javascript" onClick="return setcredit_onclick()">
        <input type=submit name=viewcredit id=viewcredit value="�鿴�����ʻ�"  language="javascript" onClick="return viewcredit_onclick()">
      </td>
      <td width="95">[ <a href="NewAgent.asp">���������</a> ]</td>
    </tr>
  </table>
  </form>
<%set rsAgentlist=nothing
set objCredit=nothing%>
</body>
</html>
