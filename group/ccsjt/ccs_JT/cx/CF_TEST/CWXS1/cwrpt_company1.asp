<%@ Language=VBScript %>
<%
'   OledbStr_cf = "provider=sqloledb;server=10.254.0.102;database=szxcw;uid=sa;pwd=123456;"  
OledbStr_cf = "provider=sqloledb;server=10.254.0.102;database=cwszx;uid=sa;pwd=123456;"  
   Set objConn_cf = Server.CreateObject("ADODB.Connection")
   objConn_cf.Open OledbStr_cf
   Set objRst=server.CreateObject ("ADODB.Recordset")
   objRst.LockType=3
   objRst.CursorType=3
   set objRst.activeConnection=objConn_cf    
   '�������ݿ�
   
   
   bdate=Request.QueryString("bdate")
   edate=Request.QueryString("edate")
   agentname=trim(Request.QueryString("ag"))
   company=Request.QueryString("com")
   depcity=Request.QueryString("dep")
   arrcity=Request.QueryString("arr")
   
   

if trim(agentname)="���д�����" then 
   sqlag=""
else
   sqlag=" and agentname='"& trim(agentname) &"'"
end if


if trim(company)="���й�˾" then 
   sqlco=""
elseif trim(company)="�⹫˾" then 
   sqlco=" and company<>'SZX' "
else
   sqlco=" and company='"& trim(company) &"'"
end if


if trim(depcity)="���к�վ" then 
   sqldep=""
else
   sqldep=" and depcity='"& trim(depcity) &"'"
end if



if trim(arrcity)="���к�վ" then 
   sqlarr=""
else
   sqlarr=" and arrcity='"& trim(arrcity) &"'"
end if

   
   
sqldate=" flightdate>='"& bdate &"' and flightdate<='"& edate &"' "

sqlwhere=sqldate+sqlco+sqlag+sqldep+sqlarr



'Response.Write sqlwhere
   

   
   
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=bdate%>��<%=edate%>�����������н�����Ʊ���</title>
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
<p align="center"><b><font size="5" color="#000099"><font color=red> <%=agentname%></font>����<font color=red><%=company%></font>�н�����Ʊ���</b>(�������ڣ�<font color=red><%=bdate%></font>��<font color=red><%=edate%></font>)</font></p>
<div align="center">
  <p align="left">  
  <center>
    <%
   
    'SqlIns ="exec cf_proc_temp '"& cstr(bdate) &"', '"& cstr(edate) &"'"
    
    SqlIns ="select ���˹�˾=company,�����=flightno, "
    SqlIns =SqlIns+"��Ʊ����=count(price),���=sum(price), "
    SqlIns =SqlIns+"��������1=ar1,��������2=ar2,��������3=ar3,������=sum(agentfee) "
    SqlIns =SqlIns+" from ticketinfo "
    SqlIns =SqlIns+" where "
    SqlIns =SqlIns+sqlwhere
    SqlIns =SqlIns+"  and (ar1<>0 or ar2<>0 or ar3<>0) "
    SqlIns =SqlIns+" group by flightno,company,ar1,ar2,ar3 "
    SqlIns =SqlIns+"order by company,flightno "
    
    objrst.Source =sqlins
    
    objrst.Open 
    
    if not (objrst.EOF and objrst.BOF) then
    
    %>  
 <table border="0" cellspacing="1" width="96%" height="1" bgcolor="#0000FF" bordercolor="#0000FF">
    <tr>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">���˹�˾</font></b></td>      
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">�����</font></b></td>      
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">��Ʊ����</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">���۶�</font></b></td>            
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">��������1</font></b></td>            
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">��������2</font></b></td>            
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">��������3</font></b></td>            
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">������</font></b></td>
    </tr>

    
    
    <%

    
    objrst.MoveFirst 
    while not objrst.EOF 
    %>
    <tr>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=trim(objrst(0))%></font></td>      
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=objrst(1)%></font></td>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=objrst(2)%></font></td>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=objrst(3)%></font></td>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=objrst(4)%></font></td>      
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=objrst(5)%></font></td> 
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=objrst(6)%></font></td> 
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=objrst(7)%></font></td> 
    </tr>
    <%
    objrst.MoveNext 
    wend
    

  %> 

  </table>
  
  <% else %>
  
  <p align="center"><b><font size="6" color="red">�޷������������ݣ�</font></b></p>
<p align="center">��</p>
<p align="center"><b><a href="./cwxs_index.asp"><font size="4" color="#0000CC">�� ��</font></a></b></p>  
  
  <%
  end if
      objrst.Close  %>
  </center>
</div>


</body>
</HTML>
