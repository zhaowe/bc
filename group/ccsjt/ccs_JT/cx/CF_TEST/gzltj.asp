<%@ Language=VBScript %>
<%
   OledbStr_cf = "provider=sqloledb;server=10.254.0.102;database=cwszx;uid=sa;pwd=123456;"  
   'OledbStr_cf = "provider=sqloledb;server=10.254.0.102;database=cwszx;uid=sa;pwd=123456;"  
   Set objConn_cf = Server.CreateObject("ADODB.Connection")
   objConn_cf.Open OledbStr_cf
   Set objRst=server.CreateObject ("ADODB.Recordset")
   objRst.LockType=3
   objRst.CursorType=3
   set objRst.activeConnection=objConn_cf    
   '�������ݿ�
   
seldatestr1=Request.QueryString("seldatestr")
lenstr=len(seldatestr1)

 pos=instr(1,seldatestr1,"|") 
 
 seldate=left(seldatestr1,pos-1)
 
   
   SelYear = cstr(year(seldate))
   SelMonth=cstr(month(seldate))
   SelDay = cstr(day(seldate))
   
   seldate1=selyear+"-"+selmonth+"-"+SelDay
   seldate2=right(seldatestr1,lenstr-pos)
   
   bdate=seldate1
   edate=seldate2
   
   
   
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=seldate1%>��<%=seldate2%>¼ƱԱ�������ͳ��</title>
<style type="text/css">


A {
	FONT-FAMILY: ����; FONT-SIZE: 15px; TEXT-DECORATION: none;color:#0000FF
}
A:hover {
	FONT-FAMILY: ����; FONT-SIZE: 15px; TEXT-DECORATION: underline; color:#FF0000
}
TD {
    FONT-FAMILY: ����; FONT-SIZE: 14px
}
</style>
</HEAD>

<body>
<p align="center"><b><font size="5" color="#000099"><%=seldate1%>��<%=seldate2%>¼ƱԱ�������ͳ��</font></b></p>
<div align="center">
  <p align="left">  
  <center>
  
 <table border="0" cellspacing="1" width="70%" height="1" bgcolor="#0000FF" bordercolor="#0000FF">
    <tr>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">¼ƱԱ</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">¼������</font></b></td>      
    </tr>
    <%
   
   
    SqlIns ="exec gzltj '"& cstr(bdate) &"', '"& cstr(edate) &"'"
    objrst.Source =sqlins
    
    objrst.Open 
    objrst.MoveFirst 
    while not objrst.EOF 
    %>
    <tr>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=objrst("¼��Ա")%></font></td>      
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=objrst("¼������")%></font></td>      
    </tr>
    <%
    objrst.MoveNext 
    wend
    objrst.Close 
    
    
    SqlIns ="select count(ticketno) from ticketinfo where flightdate>= '"& cstr(bdate) &"' and  flightdate<= '"& cstr(edate) &"' "
    objrst.Source =sqlins
    
    objrst.Open     
  %> 
    <tr>
      <td  height="1" bgcolor="#FFFFFF"><font color=red><B>�ϼ�����</B></font></td>      
      <td  height="1" bgcolor="#FFFFFF"><font color=red><B><%=objrst(0)%></B></font></td>      
    </tr>  
<% objrst.Close %>
  </table>
  </center>
</div>


</body>
</HTML>
