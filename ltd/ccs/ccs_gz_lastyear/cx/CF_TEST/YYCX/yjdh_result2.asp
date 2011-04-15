<%@ Language=VBScript %>  
  <!-- #include virtual="newinfoweb/sharecode/DataLink102.asp"-->
<%


   'OledbStr_cf = "provider=sqloledb;server=10.254.0.102;database=cwszx;uid=sa;pwd=123456;"  
   Set objConn_cf = Server.CreateObject("ADODB.Connection")
   objConn_cf.Open OledbStr_cwxs
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
<BODY>
<%

datestr=Request.Form ("datestr")




if len(trim(datestr))<>6 then

  Response.Write "输入日期数据格式错误（YYYYMM）！"
else

bdate=left(trim(datestr),4) & "-" & right(trim(datestr),2) & "-" & "01" 
edate=dateadd("d",-1,dateadd("m",1,bdate))
 
%>
 
 
   <center>
 <form method="post" action="yjdh_tj.asp"> 
 <table border="0" cellspacing="1" width="96%" height="1" bgcolor="#0000FF" bordercolor="#0000FF">
    <tr>
      
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">开始日期</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">结束日期</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">代理人</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">运价代号</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">折扣代号</font></b></td>      

    </tr>
    <%
   

sqlins="select distinct 运价代号 from ticketinfo_" & trim(datestr) & " as t "
sqlins=sqlins+" where  T.票证类型<>'REF' "
sqlins=sqlins+" and T.旅客类型<>'I' "
sqlins=sqlins+" and T.[净额（人民币）]<>0 "
sqlins=sqlins+" and T.票证种类='T' "
sqlins=sqlins+" and T.出票公司代码='CZ'"
sqlins=sqlins+" and T.始发站代码='SZX' "
    
    objrst.Source =sqlins
        
    objrst.Open 
    objrst.MoveFirst 
    i=1
    while not objrst.EOF 
    %>
    <tr>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=bdate %></font></td>  
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=edate %></font></td>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF">所有</font></td>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=objrst(0)%><INPUT type="hidden" id=zkdm name=<%="zkdm"+cstr(i)%> value="<%=objrst(0)%>"></font></td>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"> <INPUT id=zkdh name=<%="zkdh"+cstr(i)%> > </font></td>
      

    </tr>
    <%
    objrst.MoveNext 
    i=i+1
    wend
    objrst.Close 
    
    

  
sqlins=" select distinct 运价代号 from ticketinfo_" & trim(datestr) & " as t "
sqlins=sqlins+" where  T.票证类型<>'REF' "
sqlins=sqlins+" and T.旅客类型<>'I' "
sqlins=sqlins+" and T.[净额（人民币）]<>0 "
sqlins=sqlins+" and T.票证种类='T' "
sqlins=sqlins+" and T.出票公司代码='CZ'"
sqlins=sqlins+" and T.到达站代码='SZX'"
sqlins=sqlins+" and T.承运分公司航站='SZX' "
sqlins=sqlins+" and T.运价代号 not in ( select distinct 运价代号 from ticketinfo_" & trim(datestr) & " as t "
sqlins=sqlins+" where  T.票证类型<>'REF' "
sqlins=sqlins+" and T.旅客类型<>'I' "
sqlins=sqlins+" and T.[净额（人民币）]<>0 "
sqlins=sqlins+" and T.票证种类='T' "
sqlins=sqlins+" and T.出票公司代码='CZ'"
sqlins=sqlins+" and T.始发站代码='SZX') "
    
    objrst.Source =sqlins
        
    objrst.Open 
    objrst.MoveFirst 
    while not objrst.EOF 
    %>
    <tr>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=bdate %></font></td>      
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=edate %></font></td>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF">所有</font></td>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=objrst(0)%><INPUT  type="hidden" id=zkdm name=<%="zkdm"+cstr(i)%> value="<%=objrst(0)%>"></font></td>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"> <INPUT id=zkdh name=<%="zkdh"+cstr(i)%> > </font></td>
      

    </tr>
    <%
    objrst.MoveNext 
    i=i+1
    wend
    objrst.Close 
    
    

  
%>
  </table>
  
  </center>
 
 
 
 
 
    
    <%

end if

%>
<P align="center">
<INPUT  type="hidden" id=zkdm name=edate value="<%=edate%>">
<INPUT  type="hidden" id=sumid name=sumid value="<%=i%>">
<INPUT type="submit" value="提 交" id=button1 name=button1>&nbsp;</P>
</form>
</BODY>
</HTML>
