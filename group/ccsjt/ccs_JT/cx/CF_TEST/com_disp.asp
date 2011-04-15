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
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">


A {
	FONT-FAMILY: 宋体; FONT-SIZE: 15px; TEXT-DECORATION: none;color:#0000FF
}
A:hover {
	FONT-FAMILY: 宋体; FONT-SIZE: 15px; TEXT-DECORATION: underline; color:#FF0000
}
TD {
    FONT-FAMILY: 宋体; FONT-SIZE: 14px
}
</style>
</HEAD>

<body>
<p align="center"><b><font size="6" color="#000099">通  讯  录</font></b></p>
<div align="center">
  <p align="left">  
  <center>
  <a href="com_index.asp">新增名单</a><br>
 <table border="0" cellspacing="1" width="96%" height="1" bgcolor="#0000FF" bordercolor="#0000FF">
    <tr>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">姓名</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">办公电话</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">家庭电话</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">宿舍电话</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">手机</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">传呼</font></b></td>      
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">所在地</font></b></td>
    </tr>
    <%
   
    SqlIns = SqlIns + "select id,name,"
    SqlIns = SqlIns + "office_tel,home_tel,dorm_tel,mobile_tel,"
    SqlIns = SqlIns + "BP_call,email,QQ_code,city "    
    SqlIns = SqlIns + "from cf_comm order by name "    
    objrst.Source =sqlins
    
    objrst.Open 
    objrst.MoveFirst 
    while not objrst.EOF 
    %>
    <tr>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><a href="detail.asp?Q=<%=objrst("id")%>"><%=objrst("name")%></a></font></td>      
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=objrst("office_tel")%></font></td>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=objrst("home_tel")%></font></td>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=objrst("dorm_tel")%></font></td>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=objrst("mobile_tel")%></font></td>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=objrst("BP_call")%></font></td>      
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=objrst("city")%></font></td>
    </tr>
    <%
    objrst.MoveNext 
    wend
    objrst.Close 
  %> 

  </table>
  </center>
</div>


</body>
</HTML>
