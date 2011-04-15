<%@ Language=VBScript %>
<%
   OledbStr_cf = "provider=sqloledb;server=10.254.0.46;database=cftest;uid=sa;pwd=;"  
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

</HEAD>
<style type="text/css">

TD {
	FONT-FAMILY: 宋体; FONT-SIZE: 13px
}

TABLE {
	FONT-FAMILY: 宋体; FONT-SIZE: 13px
}

</style>
<body>


<div id="head" style="HEIGHT: 60px; LEFT: 3px; POSITION: absolute; TOP: 10px; WIDTH: 747px; Z-INDEX: 1">
<IMG height=59 src="../../images/airline.GIF" width=764> 
</div>
<div id="head" style="HEIGHT: 60px; LEFT: 3px; POSITION: absolute; TOP: 100px; WIDTH: 747px; Z-INDEX: 1">
<h1 align="center"><b><font color="#9900cc" face="长城新魏碑体">深圳进出港航班查询系统</font></b></h1>
<form name="tijiao" method="post" action="disp.asp">
  <div align="center">
    <center>
    <table border="1" cellspacing="1" width="92%" bordercolor="#6666ff" height="106">
      <tr>
      <%
        objrst.Source ="select distinct date from sale order by date desc"
        objrst.Open       
        objrst.MoveFirst 
      %>
        <td width="20%" height="23"><B>日期</B>
        <select size="1" name="date">
         <OPTION selected value="<%=objrst("date") %>"><%=objrst("date") %></OPTION>
         <%objrst.MoveNext  
          while not objrst.EOF %>
         <OPTION value="<%=objrst("date") %>"><%=objrst("date") %></OPTION>
         <%objrst.MoveNext  
           wend 
           objrst.Close    
           
   
           objrst.Source ="select distinct airline from sale order by airline "
           objrst.Open            
         %>
       
          </select></td>        
        </td>
        <td width="20%" height="23"><B>航空公司</B>
        <select size="1" name="airline">
         <OPTION selected value="1">全部</OPTION>
         <% while not objrst.EOF %>
         <OPTION value="<%=objrst("airline") %>"><%=objrst("airline") %></OPTION>
         <%objrst.MoveNext  
           wend 
           objrst.Close    
         %>
       
          </select></td>
        <%
          objrst.Source ="select distinct depcity from sale order by depcity "
          objrst.Open          
          objrst.MoveFirst  %>       
        <td width="20%" height="23"><B>航班号</B><input name="flightno" size="8" value="全部"></td>
        <td width="20%" height="23"><B>起飞城市</B>        
        <select size="1" name="depcity">
          <OPTION selected value="SZX">SZX</OPTION>
          <OPTION value="1">全部</OPTION>
          <% while not objrst.EOF %>
        <OPTION  value="<%=objrst("depcity") %>"><%=objrst("depcity") %></OPTION>
         <%objrst.MoveNext  
           wend 
           objrst.Close    
         %>        
          </select></td>
        <td width="20%" height="23"><B>到达城市</B>
        <select size="1" name="arrcity">
        <%
          objrst.Source ="select distinct arrcity from sale order by arrcity "
          objrst.Open           
          objrst.MoveFirst  
        %>
          <OPTION selected value="1">全部</OPTION>
          <OPTION value="SZX">SZX</OPTION>
        <%  while not objrst.EOF %>
        <OPTION  value="<%=objrst("arrcity") %>"><%=objrst("arrcity") %></OPTION>
         <%objrst.MoveNext  
           wend 
           objrst.Close    
         %> 
      </tr>
      <tr>
        <td width="103%" colspan="5" height="16"></td>
      </tr>
      <tr>
        <td width="103%" colspan="5" height="49">
          <p align="center"><input type="submit" value="提  交" name="B1"></p></td>
      </tr>
    </table>
    <p></p>
    <table>
     <tr><td>
     <b><font color="#FF0000" size="5">提示:</font></b><font color="#9D6FFF" size="4">系统提供当天所有深圳进出港航班查询<br>&nbsp&nbsp&nbsp&nbsp 起飞城市和到达城市中必须有一个为&quot;SZX&quot;</font>
     </td></tr>
    </table>
    </center>
  </div>
  <p align="center">　</p>
</form>
<p>　</p>
</div>
</body>

</HTML>
