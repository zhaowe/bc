<!--#include file="Public.inc"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>功能管理</title>
<link rel="stylesheet" href="../../style.css">
</head>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Sub btnQuery_onclick
	frminfo.action = "functioninfo.asp?tiaojian=" & frminfo.tiaojian.value & "&textfield=" & frminfo.textfield.value
	frminfo.submit  
End Sub
-->
</SCRIPT>
<body bgcolor="#FFFFFF">
<font size=6 color=blue><strong><font size="5">功能管理</font></strong></font> 
<form name="frminfo" method="post" action="functioninfo.asp" >
  <table width="610" border="0" cellspacing="1" cellpadding="4" height="15" bgcolor="#006699" >
    <tr>
              
      <td width="432"><font color="#FFFFFF">查询条件：</font> 
        <select name="tiaojian">
          <option value="functionname">功能名</option>
          <option value="description">描述</option>
          <option value="locale">版本</option>
        </select>
                
        <input type="text" name="textfield" size="20">
                <input type="button" name="BtnQuery" id=BtnQuery  value="开始查询">
                
        <input type="reset" name="Reset" value="重填">
              </td>
      <td width="157">&nbsp;</td>  
            </tr> 
          </table>
        </form>   
<%  
dim tiaojian
dim textfield  
    tiaojian=Request.QueryString ("tiaojian")  
    textfield=Request.QueryString ("textfield")  
    textfield=trim(textfield)  
dim funinfo
dim obj
on error resume next
set obj=server.createobject("com_usermanage1.clsFunction1")
set funinfo=server.CreateObject("ADODB.Recordset")
set funinfo=Obj.SearchFunctionInfo(textfield,tiaojian)
if err.number<>0 then  
   Ierror=err.number  
   err.clear  
set objdml=nothing  
   'Response.End   
   response.redirect "../../Sorry.asp?Errorno=" &ierror  
end if   
  %>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>  
<!--  
  
function button1_onclick() {  
	form1.text.value="edit"  
	form1.submit()  
	return true   
}  
  
function button2_onclick() {  
	form1.text.value="del"  
	form1.submit()  
	return true   
}  
  
//-->  
</SCRIPT> 
<%const PAGE_SIZE = 8
funInfo.PageSize = PAGE_SIZE
Dim iCurrentPage

if CInt(Request.QueryString("PageNo"))>=1 and CInt(Request.QueryString("PageNo"))<=funInfo.PageCount then
	iCurrentPage = CInt(Request.QueryString("PageNo"))
else
	iCurrentPage =1
end if%>
<%If not funInfo.EOF Then
	funInfo.AbsolutePage = iCurrentPage
	If iCurrentPage > 1 Then
		Response.Write  "<A href='functioninfo.asp?PageNo=" & 1 & "& functionid=" & Request.QueryString("functionid") &  "'>首页</a>&nbsp;&nbsp;"
		Response.Write  "<A href='functioninfo.asp?PageNo=" & (iCurrentPage-1) & "& functionid=" & Request.QueryString("functionid") &  "'>上一页</a>&nbsp;&nbsp;"
	End If
	If iCurrentPage < funInfo.PageCount Then
		Response.Write "<A href='functioninfo.asp?PageNo=" & (iCurrentPage+1) & "&groupid=" & Request.QueryString("functionid") &  "'>下一页</a>&nbsp;&nbsp;"
	   Response.Write  "<A href='functioninfo.asp?PageNo=" &(funinfo.pagecount)& "& functionid=" & Request.QueryString("functionid") &  "'>尾页</a>&nbsp;&nbsp;"

	End If
%> 第<font color=red> <%=iCurrentPage%></font>/<%=funInfo.PageCount%> 页<br> 

<FORM action="bosom.asp" method="post" name="form1">      
      
  <table border=0 cellPadding=4  cellSpacing=1 width="610" bgcolor="#000000">
    <tr  bgcolor="#003333"> 
      <td valign=top><font color="#FFFFFF">操作</font></td>
      <td valign=top><font color="#FFFFFF">功能名</font></td>
      <td valign=top><font color="#FFFFFF">描述</font></td>
      <td valign=top><font color="#FFFFFF">序号</font></td>
      <td valign=top><font color="#FFFFFF">父功能名</font></td>
      <td valign=top><font color="#FFFFFF">功能类型 </font></td>
    </tr>
    <%for i=1 to PAGE_SIZE%>  
    <tr> 
       <td  bgcolor="#FFFFFF"><INPUT id=radio1 name=functionid type=radio value="<%=funinfo(0)%>" <%if  i mod 8 =1 then Response.Write  "checked" end if%>></td>
      <td align=top height="20" bgcolor="#FFFFFF" ><%=funinfo("functionname")%> </td>
      <td align=top height="20" bgcolor="#FFFFFF" ><%=funinfo("description")%></td>
      <td align=top height="20" bgcolor="#FFFFFF" > <%=funinfo("ordernum")%></td>
      <td align=top height="20" bgcolor="#FFFFFF" > <%=funinfo("ffunctionname")%></td>
      <td bgcolor="#ffFFFF">
      <%f=trim(funinfo("functiontype"))%>
      <%select case f
       case "M"
       Response.Write "菜单"
       case "F"
       Response.Write "功能"  
       case "P"
       Response.Write "页面" 
       case "MF" 
       Response.Write "菜单,功能"
       case "MP"
       Response.Write "菜单,页面"
       case "FP"
       Response.Write "功能,页面"
       case "MFP" 
       Response.Write "功能,菜单,页面"
       End Select
      %>
     
       </td>
     </tr>
    <%   
		funinfo.movenext
		If funInfo.EOF Then 
			Exit For 
		End If 
	next
 
%> 
  </TABLE> 
   <%
Else 
	Response.Write "当前没有记录" 
End If %>
<br>
  <table width="610" border="0" cellpadding="4" cellspacing="1">
    <tr>
      <td width="281"> 
        <input name=text type=hidden>
        <input name=button1 type=button value=修改 language=javascript onClick="return button1_onclick()">
        <input name=button2 type=button value=删除 language=javascript onClick="return button2_onclick()">
      </td>
      <td width="308" align="center" bordercolor="#FFFFFF">[ <a href="addfunction.asp">增加功能</a> 
        ]</td>
    </tr>
  </table>
</FORM>
<%set obj=nothing%>
</body>  
</html>  
  
  
