<%@ Language=VBScript %>
<%
   OledbStr_cf = "provider=sqloledb;server=10.254.0.46;database=cftest;uid=sa;pwd=szx6275;"  
   Set objConn_cf = Server.CreateObject("ADODB.Connection")
   objConn_cf.Open OledbStr_cf
   Set objRst=server.CreateObject ("ADODB.Recordset")
   objRst.LockType=3
   objRst.CursorType=3
   set objRst.activeConnection=objConn_cf    
   '连接数据库
   
   id=Request.QueryString ("Q")
   
   SqlIns ="select * from cf_comm where id='"& id &"' "    
   objrst.Source =sqlins    
   objrst.Open    
   
   if objrst("sex")=0 then
      sex="女"
   else
      sex="男"
   end if   
   
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<style type="text/css">

TD {
    FONT-FAMILY: 宋体; FONT-SIZE: 15px
}
</style>
<body>
<p align="center">　
</p>

<p align="center"><b><font size="5"><font color="#CC3300"><%=objrst("name")%></font>——<font color="#0066FF">详细资料</font></font></b>
</p>

 <table border="0" cellspacing="1" width="100%" height="1" bgcolor="#0000FF" bordercolor="#0000FF">
    <tr>
      <td width="4%" height="1" bgcolor="#FFFFFF"><font color="#990099"><b>姓名</b></font></td>
      <td width="7%" height="1" bgcolor="#FFFFFF"><%=objrst("name")%></td>
      <td width="4%" height="1" bgcolor="#FFFFFF"><font color="#990099"><b>性别</b></font></td>
      <td width="3%" height="1" bgcolor="#FFFFFF"><%=sex%></td>
      <td width="4%" height="1" bgcolor="#FFFFFF"><font color="#990099"><b>年龄</b></font></td>
      <%if objrst("age")>0 then%>
      <td width="3%" height="1" bgcolor="#FFFFFF"><%=objrst("age")%></td>
      <%else%>
      <td width="3%" height="1" bgcolor="#FFFFFF"></td>
      <%end if%>
      <td width="4%" height="1" bgcolor="#FFFFFF"><font color="#990099"><b>生肖</b></font></td>
      <td width="3%" height="1" bgcolor="#FFFFFF"><%=objrst("shengxiao")%></td>
      <td width="4%" height="1" bgcolor="#FFFFFF"><font color="#990099"><b>星座</b></font></td>
      <td width="6%" height="1" bgcolor="#FFFFFF"><%=objrst("constellation")%></td>
      <td width="4%" height="1" bgcolor="#FFFFFF"><font color="#990099"><b>血型</b></font></td>
      <td width="3%" height="1" bgcolor="#FFFFFF"><%=objrst("bloodtype")%></td>      
      <td width="8%" height="1" bgcolor="#FFFFFF"><font color="#990099"><b>出生日期</b></font></td>
      <%if objrst("age")>0 then%>
      <td width="15%" height="1" bgcolor="#FFFFFF"><%=objrst("birthday")%></td>
      <%else%>
      <td width="15%" height="1" bgcolor="#FFFFFF"></td>
      <%end if%>
    </tr>
 </table>  
 <table border="0" cellspacing="1" width="100%" height="1" bgcolor="#0000FF" bordercolor="#0000FF">
    <tr>
      <td width="8%" height="1" bgcolor="#FFFFFF"><font color="#990099"><b>办公电话</b></font></td>
      <td width="15%" height="1" bgcolor="#FFFFFF"><%=objrst("office_tel")%></td>
      <td width="8%" height="1" bgcolor="#FFFFFF"><font color="#990099"><b>家庭电话</b></font></td>
      <td width="16%" height="1" bgcolor="#FFFFFF"><%=objrst("home_tel")%></td>
      <td width="8%" height="1" bgcolor="#FFFFFF"><font color="#990099"><b>宿舍电话</b></font></td>
      <td width="17%" height="1" bgcolor="#FFFFFF"><%=objrst("dorm_tel")%></td>
    </tr>
 </table>  
 <table border="0" cellspacing="1" width="100%" height="1" bgcolor="#0000FF" bordercolor="#0000FF">
    <tr>
      <td width="5%" height="1" bgcolor="#FFFFFF"><font color="#990099"><b>手机</b></font></td>
      <td width="9%" height="1" bgcolor="#FFFFFF"><%=objrst("mobile_tel")%></td>
      <td width="5%" height="1" bgcolor="#FFFFFF"><font color="#990099"><b>传呼</b></font></td>
      <td width="15%" height="1" bgcolor="#FFFFFF"><%=objrst("bp_call")%></td>
      <td width="6%" height="1" bgcolor="#FFFFFF"><font color="#990099"><b>EMAIL</b></font></td>
      <td width="26%" height="1" bgcolor="#FFFFFF"><a href="mailto:<%=objrst("email")%>"><%=objrst("email")%></a></td>
      <td width="3%" height="1" bgcolor="#FFFFFF"><font color="#990099"><b>QQ</b></font></td>
      <td width="12%" height="1" bgcolor="#FFFFFF"><%=objrst("qq_code")%></td>
    </tr>
 </table>   
 <table border="0" cellspacing="1" width="100%" height="1" bgcolor="#0000FF" bordercolor="#0000FF">
    <tr>
      <td width="8%" height="1" bgcolor="#FFFFFF"><font color="#990099"><b>工作单位</b></font></td>
      <td width="73%" height="1" bgcolor="#FFFFFF"><%=objrst("corporation")%></td>
    </tr>
    <tr>
      <td width="8%" height="1" bgcolor="#FFFFFF"><font color="#990099"><b>公司地址</b></font></td>
      <td width="73%" height="1" bgcolor="#FFFFFF"><%=objrst("office_addr")%></td>
    </tr>
    <tr>
      <td width="8%" height="1" bgcolor="#FFFFFF"><font color="#990099"><b>家庭地址</b></font></td>
      <td width="73%" height="1" bgcolor="#FFFFFF"><%=objrst("home_addr")%></td>
    </tr>
    <tr>
      <td width="8%" height="1" bgcolor="#FFFFFF"><font color="#990099"><b>公寓地址</b></font></td>
      <td width="73%" height="1" bgcolor="#FFFFFF"><%=objrst("dorm_addr")%></td>
    </tr>            
 </table>   
 <table border="0" cellspacing="1" width="100%" height="1" bgcolor="#0000FF" bordercolor="#0000FF">
    <tr>
      <td width="8%" height="1" bgcolor="#FFFFFF"><font color="#990099">&nbsp;&nbsp; 
        <b>备注</b></font></td>
      <td width="73%" height="1" bgcolor="#FFFFFF"><%=objrst("memo")%></td>
    </tr>         
 </table>    
<p align="center"><b><font size="4"><font color="#0066FF"><a href="com_modi.asp?Q=<%=id%>">修  改</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="com_disp.asp">返  回</a></font></font></b>
</p> 
</body>
</HTML>
<%
objrst.Close 

%>
