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
if Err.number<>0 then  '加盟点类别集
   ErrNo=Err.number 
   Err.Clear 
   set objCredit=nothing
   set rsAgentType=nothing
   Response.Redirect("../../sorry.asp?ErrorNo=" + cstr(ErrNo))
end if


Set rsAgentList = objCredit.GetAllAgentInfo("V",Application("UseObject"),Application("Locale"))
if Err.number<>0 then  '上级节点集和
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
			alert("请在代理点Office栏内输入代理点编号!")
			return false
		}	
if(document.form1.text2.value == ""){
			alert("请在代理点名称栏内输入该代理点全称!")
			return false
		}	
if(document.form1.text3.value == ""){
			alert("请在代理点简称栏内输入该代理点简称!")
			return false
		}	
		
if(document.form1.text5.value == ""){
			alert("请在联系地址栏内输入该代理点单位联系地址!")
			return false
		}	
if(document.form1.text6.value == ""){
			alert("请在联系人栏内输入该代理点单位联系人!")
			return false
		}	
if(document.form1.text7.value == ""){
			alert("请在联系电话栏内输入该代理点单位联系电话!")
			return false
		}	
if(document.form1.text8.value == ""){
			alert("请在所在城市栏内输入该代理点所在城市!")
			return false
		}	
if(document.form1.text13.value == ""){
			alert("请在操作者栏内输入办理该代理点入户的操作员!")
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
  <p><b><font color="#0000FF">新增代理点(资料输入)</font></b></p>
  <TABLE border=0 cellPadding=4 cellSpacing=1 width="610" bgcolor="#000000">
    <TR> 
      <TD style="WIDTH: 25%" width="25%" bgcolor="#003333"><font color="#FFFFFF">代理点Office:</font></TD>
      <TD style="WIDTH: 25%" width="25%" bgcolor="#FFFFFF"> 
        <INPUT id=text1 
      name=AgentOffice style="HEIGHT: 22px; WIDTH: 88px" >
      </TD>
      <TD style="WIDTH: 25%" width="25%" bgcolor="#003333"><font color="#FFFFFF">代理点类别</font></TD>
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
      <TD bgcolor="#003333"><font color="#FFFFFF">代理点名称</font></TD>
      <TD bgcolor="#FFFFFF"> 
        <INPUT id=text2 name=AgentName
      style="HEIGHT: 22px; WIDTH: 88px" >
      </TD>
      <TD bgcolor="#003333"><font color="#FFFFFF">代理点简称</font></TD>
      <TD bgcolor="#FFFFFF"> 
        <INPUT id=text3 name=AgentShortName
      style="HEIGHT: 22px; WIDTH: 88px" >
      </TD>
    </TR>
    <TR> 
      <TD bgcolor="#003333"><font color="#FFFFFF">上级代理点</font></TD>
      <TD bgcolor="#FFFFFF"> 
        <SELECT id=select1 name=FAgentID size=1  style="HEIGHT: 29px; WIDTH: 76px">
          <% for i=1 to rsAgentList.RecordCount %> 
          <Option <%if rsAgentList("AgentType")="H" then%> selected <%end if%>  value="<%=rsAgentList("AgentID")%>"> 
          <%=rsAgentList("AgentName")%></OPTION>
          <%rsAgentList.MoveNext 
    next%> 
        </SELECT>
      </TD>
      <TD bgcolor="#003333"><font color="#FFFFFF">联系地址</font></TD>
      <TD bgcolor="#FFFFFF"> 
        <INPUT id=text5 name=lxrAdd 
      style="HEIGHT: 22px; WIDTH: 88px" >
      </TD>
    </TR>
    <TR> 
      <TD bgcolor="#003333"><font color="#FFFFFF">联系人</font></TD>
      <TD bgcolor="#FFFFFF"> 
        <INPUT id=text6 name=lxrName 
      style="HEIGHT: 22px; WIDTH: 88px" >
      </TD>
      <TD bgcolor="#003333"><font color="#FFFFFF">联系电话</font></TD>
      <TD bgcolor="#FFFFFF"> 
        <INPUT id=text7 name=lxrPhone
      style="HEIGHT: 22px; WIDTH: 88px" >
      </TD>
    </TR>
    <TR> 
      <TD bgcolor="#003333"><font color="#FFFFFF">所在城市</font></TD>
      <TD bgcolor="#FFFFFF"> 
        <INPUT id=text8 name=AgentCity
      style="HEIGHT: 22px; WIDTH: 88px" >
      </TD>
      <TD bgcolor="#003333"><font color="#FFFFFF">开户银行</font></TD>
      <TD bgcolor="#FFFFFF"> 
        <INPUT id=text9 name=OpenBank 
      style="HEIGHT: 22px; WIDTH: 88px" >
      </TD>
    </TR>
    <TR> 
      <TD bgcolor="#003333"><font color="#FFFFFF">开户帐号</font></TD>
      <TD bgcolor="#FFFFFF"> 
        <INPUT id=text10 name=OpenAccount 
      style="HEIGHT: 22px; WIDTH: 88px" >
      </TD>
      <TD bgcolor="#003333"><font color="#FFFFFF">协议代号</font></TD>
      <TD bgcolor="#FFFFFF"> 
        <INPUT id=text11 name=ProtocalNo
      style="HEIGHT: 22px; WIDTH: 88px" >
      </TD>
    </TR>
    <TR> 
      <TD bgcolor="#003333"><font color="#FFFFFF">协议日期</font></TD>
      <TD bgcolor="#FFFFFF"> 
        <INPUT id=text12 name=ProtocalDate
      style="HEIGHT: 22px; WIDTH: 88px" value=<%=Today%>>
      </TD>
      <TD bgcolor="#003333"><font color="#FFFFFF">虚实选择</font></TD>
      <TD bgcolor="#FFFFFF"> 
        <SELECT id=select2 name=AgentEntity 
      style="HEIGHT: 22px; WIDTH: 77px">
          <OPTION   selected  value="T">实点</OPTION>
          <OPTION   value="F">虚点</OPTION>
        </SELECT>
      </TD>
    </TR>
  </TABLE>
  <p>
    <input type="submit" value="新增代理点" id=submit1 name=submit1 onClick="return submit1_Click()">
     <input type="button" value="取消" id=submit1 name=submit12 onClick="self.history.back()">
  </p>
</FORM>

</BODY>
</HTML>
