<%@ Language=VBScript %>  <!-- #include virtual="sharecode/DataLink102.asp"-->
<%



Function DouStr(Smonth) 

     If Len(Trim(Smonth)) <> 2 Then
        DouStr = "0" + Trim(Smonth)
     Else
        DouStr = Trim(Smonth)
     End If
End Function
%>

<%
'   'OledbStr_cwxs = "provider=sqloledb;server=10.254.0.102;database=szxcw;uid=sa;pwd=123456;"  
'OledbStr_cwxs = "provider=sqloledb;server=10.254.0.102;database=cwszx;uid=sa;pwd=123456;"  
   Set objConn_cf = Server.CreateObject("ADODB.Connection")
   objConn_cf.Open OledbStr_cwxs
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

'dim jdb as string

'Response.Write sqlwhere

JDB = " ticketinfo_" + CStr(Year(cdate(edate))) + "_" + DouStr(CStr(Month(cdate(edate)))) + " "

   
if cdate(edate)>=cdate("2005-10-01") and cdate(edate)<=cdate("2005-12-31") then
   jdb=" ticketinfo_2005d4jd  "
end if   
   
if cdate(edate)>=cdate("2005-07-01") and cdate(edate)<=cdate("2005-9-30") then
   jdb=" ticketinfo_2005d3jd  "
end if      

if cdate(edate)>=cdate("2005-04-01") and cdate(edate)<=cdate("2005-06-30") then
   jdb=" ticketinfo_2005d2jd  "
end if   
   
if cdate(edate)>=cdate("2005-02-01") and cdate(edate)<=cdate("2005-03-31") then
   jdb=" ticketinfo_2005d1jd  "
end if  
   
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=bdate%>��<%=edate%>�����������н�����Ʊ��ϸ��</title>
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
<p align="center"><b><font size="5" color="#000099"><font color=red><%=agentname%></font>����<font color=red><%=company%></font>�н�����Ʊ��ϸ��</b>(�������ڣ�<font color=red><%=bdate%></font>��<font color=red><%=edate%></font>)</font></p>
<div align="center">
  <p align="left">  
  <center>
    <%
   
    'SqlIns ="exec cf_proc_temp '"& cstr(bdate) &"', '"& cstr(edate) &"'"
    
    
    SqlIns =" select ������=agentname,��Э��=agentsym,��������=flightdate,�����=flightno, "
    SqlIns =SqlIns+"��ɳ���=depcity,�������=arrcity, "
    SqlIns =SqlIns+"Ʊ��=ticketno,��λ=berthname,�˼۴���=pricecode,�ۿ۴���=berthcode,Ʊ��=price,��������=saledate,ar1=isnull(ar1,0),ar2=isnull(ar2,0),ar3=isnull(ar3,0),ar4=isnull(ar4,0),ar5=isnull(ar5,0),ar6=isnull(ar6,0) "
    SqlIns =SqlIns+" from "+ jdb
    SqlIns =SqlIns+" where "
    SqlIns =SqlIns+sqlwhere
    SqlIns =SqlIns+"  and (ar1<>0 or ar2<>0 or ar3<>0 or ar4<>0 or ar5<>0) "    
    SqlIns =SqlIns+"order by agentname,agentsym,flightdate,depcity,arrcity,flightno,price desc "
    
    objrst.Source =sqlins
    
    'Response.Write sqlins
    objrst.Open 
    
    if not (objrst.EOF and objrst.BOF) then
    
    %>  
 <table border="0" cellspacing="1" width="96%" height="1" bgcolor="#0000FF" bordercolor="#0000FF">
    <tr>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">������</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">��Э��</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">��������</font></b></td>      
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">�����</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">��ɳ���</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">�������</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">Ʊ��</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">��λ</font></b></td>

      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">Ʊ��(����)</font></b></td>            
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">��������</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">������1</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">������2</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">������3</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">������4</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">������5</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">������6</font></b></td>
    </tr>

    
    
    <%

    
    objrst.MoveFirst 
    while not objrst.EOF 
    
    if cint(objrst(12))<>0 then
      ar1="<>0"
    else
      ar1="=0"
    end if
        
    if cint(objrst(13))<>0 then
      ar2="<>0"
    else
      ar2="=0"
    end if
    
    if cint(objrst(14))<>0 then
      ar3="<>0"
    else
      ar3="=0"
    end if
    if cint(objrst(15))<>0 then
      ar4="<>0"
    else
      ar4="=0"
    end if
        
    if cint(objrst(16))<>0 then
      ar5="<>0"
    else
      ar5="=0"
    end if
    
    if cint(objrst(17))<>0 then
      ar6="<>0"
    else
      ar6="=0"
    end if    
    
    %>
    <tr>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=trim(objrst(0))%></font></td>      
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=trim(objrst(1))%></font></td>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=objrst(2)%></font></td>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=objrst(3)%></font></td>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=objrst(4)%></font></td>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=objrst(5)%></font></td>      
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=objrst(6)%></font></td>
  
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=objrst(7)%></font></td>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=objrst(10)%></font></td>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=objrst(11)%></font></td>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=ar1%></font></td>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=ar2%></font></td>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=ar3%></font></td>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=ar4%></font></td>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=ar5%></font></td>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=ar6%></font></td>      
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
