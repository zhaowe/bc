<%@ Language=VBScript %>
<!--#include file="../../include/UserCheck.asp"-->

<%

AgentID=session("AgentID")

if len(AgentID)=0 then
	ErrNo=10785
	Response.Redirect("../../sorry.asp?ErrorNo=" + cstr(ErrNo))
end if
on error resume next
'declare some variable
set rsAgentInfo=server.CreateObject("ADODB.recordset")
set rsAgentList=server.CreateObject("ADODB.recordset")
'set rsAgentType=server.CreateObject("ADODB.recordset")
set objDML=server.CreateObject("com_DML.clsDML")
set objCredit=server.CreateObject("com_Agent.clsAgent")

'get recordset from database
'sqlStr="select AgentType,AgentTypeName from AgentTypeLocale where AgentType<>'H' and locale='" & application("Locale") & "'"
'set rsAgentType=objDML.ExeSelect(sqlStr,7)
'if Err.number<>0 then  '代理点类别
''   ErrNo=Err.number 
'   Err.Clear 
'   set objCredit=nothing
'   set rsAgentType=nothing
'   Response.Redirect("../../sorry.asp?ErrorNo=" + cstr(ErrNo))
'end if


'Set rsAgentList = objCredit.GetAllAgentInfo("V",Application("UseObject"),Application("Locale"))
'if Err.number<>0 then  '上级节点集和
''   ErrNo=Err.number 
'   Err.Clear 
'   set objCredit=nothing
'   set rsAgentList=nothing
'   Response.Redirect("../../sorry.asp?ErrorNo=" + cstr(ErrNo))
'end if

Set rsAgentInfo = objCredit.GetAgentInfo(AgentID)
if Err.number<>0 then   '本节点数据
   ErrNo=Err.number 
   Err.Clear 
   set objCredit=nothing
   set rsAgentInfo=nothing
   Response.Redirect("../../sorry.asp?ErrorNo=" + cstr(ErrNo))
end if

if rsAgentInfo.RecordCount=0 then
   ErrNo=10777   '记录没找到!
   set objCredit=nothing
   set rsAgentInfo=nothing
   Response.Redirect("../../sorry.asp?ErrorNo=" + cstr(ErrNo))
end if

%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" href="../../style.css">
</HEAD>
<BODY bgcolor="#FFFFFF">
<p>
<b> 查看本代理点信息</b> 
  <TABLE border=0 cellPadding=4 cellSpacing=1 width="610" style="BORDER-top:1px solid rgb(00,00,00)" bgcolor="#000000">
    <TR> 
      <TD style="WIDTH: 25%" width="23%" align="right" bgcolor="#003333"><font color="#FFFFFF">代理点Office</font></TD>
      <TD style="WIDTH: 25%" width="22%" bgcolor="#FFFFFF"> 
        <INPUT id=text1 
      name=AgentOffice style="HEIGHT: 22px; WIDTH: 88px" value="<%=rsAgentInfo("AgentOffice")%>"  readonly>
      </TD>
      <TD style="WIDTH: 25%" width="29%" align="right" bgcolor="#003333"><font color="#FFFFFF">代理点类别</font></TD>
      <TD style="WIDTH: 25%" width="22%" bgcolor="#FFFFFF"> 
        <INPUT id=text2 name=AgentTypeName style="HEIGHT: 22px; WIDTH: 88px" value="<%=rsAgentInfo("AgentTypeName")%>"  readonly>
      </TD>
    </TR>
    <TR> 
      <TD width="23%" align="right" bgcolor="#003333"><font color="#FFFFFF">代理点名称</font></TD>
      <TD width="22%" bgcolor="#FFFFFF"> 
        <INPUT id=text3 name=AgentName readonly
      style="HEIGHT: 22px; WIDTH: 88px" value="<%=trim(rsAgentInfo("AgentName"))%>">
      </TD>
      <TD width="29%" align="right" bgcolor="#003333"><font color="#FFFFFF">代理点简称</font></TD>
      <TD width="26%" bgcolor="#FFFFFF"> 
        <INPUT id=text4 name=AgentShortName readonly
      style="HEIGHT: 22px; WIDTH: 88px" value="<%=trim(rsAgentInfo("AgentShortName"))%>">
      </TD>
    </TR>
    <TR> 
      <TD width="23%" align="right" bgcolor="#003333"><font color="#FFFFFF">上级代理点</font></TD>
      <TD width="26%" bgcolor="#FFFFFF"> 
        <INPUT id=text5 name=AgentShortName readonly
      style="HEIGHT: 22px; WIDTH: 88px" value="<%=trim(rsAgentInfo("FAgentName"))%>">
      </TD>
      <TD width="29%" align="right" bgcolor="#003333"><font color="#FFFFFF">联系地址</font></TD>
      <TD width="26%" bgcolor="#FFFFFF"> 
        <INPUT id=text6 name=lxrAdd  readonly
      style="HEIGHT: 22px; WIDTH: 88px" value="<%=trim(rsAgentInfo("lxrAdd"))%>">
      </TD>
    </TR>
    <TR> 
      <TD width="23%" align="right" bgcolor="#003333"><font color="#FFFFFF">联系人</font></TD>
      <TD width="22%" bgcolor="#FFFFFF"> 
        <INPUT id=text7 name=lxrName  readonly
      style="HEIGHT: 22px; WIDTH: 88px" value="<%=trim(rsAgentInfo("lxrName"))%>">
      </TD>
      <TD width="29%" align="right" bgcolor="#003333"><font color="#FFFFFF">联系电话</font></TD>
      <TD width="26%" bgcolor="#FFFFFF"> 
        <INPUT id=text8 name=lxrPhone readonly
      style="HEIGHT: 22px; WIDTH: 88px" value="<%=trim(rsAgentInfo("lxrPhone"))%>">
      </TD>
    </TR>
    <TR> 
      <TD width="23%" align="right" bgcolor="#003333"><font color="#FFFFFF">所在城市</font></TD>
      <TD width="22%" bgcolor="#FFFFFF"> 
        <INPUT id=text9 name=AgentCity readonly
      style="HEIGHT: 22px; WIDTH: 88px" value="<%=trim(rsAgentInfo("AgentCity"))%>">
      </TD>
      <TD width="29%" align="right" bgcolor="#003333"><font color="#FFFFFF">开户银行</font></TD>
      <TD width="26%" bgcolor="#FFFFFF"> 
        <INPUT id=text10 name=OpenBank  readonly
      style="HEIGHT: 22px; WIDTH: 88px" value="<%=trim(rsAgentInfo("OpenBank"))%>">
      </TD>
    </TR>
    <TR> 
      <TD width="23%" align="right" bgcolor="#003333"><font color="#FFFFFF">开户帐号</font></TD>
      <TD width="22%" bgcolor="#FFFFFF"> 
        <INPUT id=text10 name=OpenAccount  readonly
      style="HEIGHT: 22px; WIDTH: 88px" value="<%=trim(rsAgentInfo("OpenAccount"))%>">
      </TD>
      <TD width="29%" align="right" bgcolor="#003333"><font color="#FFFFFF">协议代号</font></TD>
      <TD width="26%" bgcolor="#FFFFFF"> 
        <INPUT id=text11 name=ProtocalNo readonly
      style="HEIGHT: 22px; WIDTH: 88px" value="<%=trim(rsAgentInfo("ProtocalNo"))%>">
      </TD>
    </TR>
    <TR> 
      <TD width="23%" align="right" bgcolor="#003333"><font color="#FFFFFF">协议日期</font></TD>
      <TD width="22%" bgcolor="#FFFFFF"> 
        <INPUT id=text12 name=ProtocalDate readonly
      style="HEIGHT: 22px; WIDTH: 88px" value="<%=trim(rsAgentInfo("ProtocalDate"))%>">
      </TD>
      <TD width="29%" align="right" bgcolor="#003333"><font color="#FFFFFF">虚实选择</font></TD>
      <TD width="26%" bgcolor="#FFFFFF"> 
        <SELECT id=select2 name=AgentEntity disabled
      style="HEIGHT: 22px; WIDTH: 77px">
          <OPTION  <%if ucase(rsAgentInfo("AgentEntity"))="T" then%> selected <%end if%> value="T">实点</OPTION>
          <OPTION  <%if ucase(rsAgentInfo("AgentEntity"))="F" then%> selected <%end if%> value="F">虚点</OPTION>
        </SELECT>
      </TD>
    </TR>
  </TABLE>
  <p> 
    <input type="button" value="返回" id=submit1 name=submit12 onClick="self.history.back()">
  </p>
</FORM>

</BODY>
<%set rsAgentInfo=nothing
set rsAgentList=nothing
set rsAgentType=nothing
set objCredit=nothing%>
</HTML>
