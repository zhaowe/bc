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
   '�������ݿ�
   


%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
<%

datestr=Request.Form ("datestr")




if len(trim(datestr))<>6 then

  Response.Write "�����������ݸ�ʽ����YYYYMM����"
else

bdate=left(trim(datestr),4) & "-" & right(trim(datestr),2) & "-" & "01" 
edate=dateadd("d",-1,dateadd("m",1,bdate))
 
%>
 
 
   <center>
 <form method="post" action="yjdh_tj.asp"> 
 <table border="0" cellspacing="1" width="96%" height="1" bgcolor="#0000FF" bordercolor="#0000FF">
    <tr>
      
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">��ʼ����</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">��������</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">������</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">�˼۴���</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">�ۿ۴���</font></b></td>      

    </tr>
    <%
   

sqlins="select distinct �˼۴��� from ticketinfo_" & trim(datestr) & " as t "
sqlins=sqlins+" where  T.Ʊ֤����<>'REF' "
sqlins=sqlins+" and T.�ÿ�����<>'I' "
sqlins=sqlins+" and T.[�������ң�]<>0 "
sqlins=sqlins+" and T.Ʊ֤����='T' "
sqlins=sqlins+" and T.��Ʊ��˾����='CZ'"
sqlins=sqlins+" and T.ʼ��վ����='SZX' "
    
    objrst.Source =sqlins
        
    objrst.Open 
    objrst.MoveFirst 
    i=1
    while not objrst.EOF 
    %>
    <tr>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=bdate %></font></td>  
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=edate %></font></td>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF">����</font></td>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=objrst(0)%><INPUT type="hidden" id=zkdm name=<%="zkdm"+cstr(i)%> value="<%=objrst(0)%>"></font></td>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"> <INPUT id=zkdh name=<%="zkdh"+cstr(i)%> > </font></td>
      

    </tr>
    <%
    objrst.MoveNext 
    i=i+1
    wend
    objrst.Close 
    
    

  
sqlins=" select distinct �˼۴��� from ticketinfo_" & trim(datestr) & " as t "
sqlins=sqlins+" where  T.Ʊ֤����<>'REF' "
sqlins=sqlins+" and T.�ÿ�����<>'I' "
sqlins=sqlins+" and T.[�������ң�]<>0 "
sqlins=sqlins+" and T.Ʊ֤����='T' "
sqlins=sqlins+" and T.��Ʊ��˾����='CZ'"
sqlins=sqlins+" and T.����վ����='SZX'"
sqlins=sqlins+" and T.���˷ֹ�˾��վ='SZX' "
sqlins=sqlins+" and T.�˼۴��� not in ( select distinct �˼۴��� from ticketinfo_" & trim(datestr) & " as t "
sqlins=sqlins+" where  T.Ʊ֤����<>'REF' "
sqlins=sqlins+" and T.�ÿ�����<>'I' "
sqlins=sqlins+" and T.[�������ң�]<>0 "
sqlins=sqlins+" and T.Ʊ֤����='T' "
sqlins=sqlins+" and T.��Ʊ��˾����='CZ'"
sqlins=sqlins+" and T.ʼ��վ����='SZX') "
    
    objrst.Source =sqlins
        
    objrst.Open 
    objrst.MoveFirst 
    while not objrst.EOF 
    %>
    <tr>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=bdate %></font></td>      
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=edate %></font></td>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF">����</font></td>
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
<INPUT type="submit" value="�� ��" id=button1 name=button1>&nbsp;</P>
</form>
</BODY>
</HTML>
