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
   '连接数据库


   Set objRst1=server.CreateObject ("ADODB.Recordset")
   objRst1.LockType=3
   objRst1.CursorType=3
   set objRst1.activeConnection=objConn_cf    
   '连接数据库   
   
   bdate=Request.QueryString("bdate")
   edate=Request.QueryString("edate")
   agentname=trim(Request.QueryString("ag"))
   company=Request.QueryString("com")
   depcity=Request.QueryString("dep")
   arrcity=Request.QueryString("arr")
   
   

if trim(agentname)="所有代理人" then 
   sqlag=""
else
   sqlag=" and agentname='"& trim(agentname) &"'"
end if


if trim(company)="所有公司" then 
   sqlco=""
elseif trim(company)="外公司" then 
   sqlco=" and company<>'SZX' "
else
   sqlco=" and company='"& trim(company) &"'"
end if


if trim(depcity)="所有航站" then 
   sqldep=""
else
   sqldep=" and depcity='"& trim(depcity) &"'"
end if



if trim(arrcity)="所有航站" then 
   sqlarr=""
else
   sqlarr=" and arrcity='"& trim(arrcity) &"'"
end if

   
   
sqldate=" flightdate>='"& bdate &"' and flightdate<='"& edate &"' "

sqlwhere=sqldate+sqlco+sqlag+sqldep+sqlarr


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


'Response.Write sqlwhere
   
    
   
    'SqlIns ="exec cf_proc_temp '"& cstr(bdate) &"', '"& cstr(edate) &"'"
   
    SqlIns ="select 代理人=agentname, "
    SqlIns =SqlIns+"奖励额1=sum(ar1*price)/100,奖励额2=sum(ar2*price)/100,奖励额3=sum(ar3*price)/100,奖励额4=sum(ar4*price)/100,奖励额5=sum(ar5*price)/100, "
    SqlIns =SqlIns+"合计= (sum(ar1*price)+sum(ar2*price)+sum(ar3*price)+sum(ar4*price)+sum(ar5*price))/100  "
    SqlIns =SqlIns+" from "+ jdb
    SqlIns =SqlIns+" where "
    SqlIns =SqlIns+sqlwhere
    SqlIns =SqlIns+"  and (ar1<>0 or ar2<>0 or ar3<>0 or ar4<>0  or ar5<>0)"
    SqlIns =SqlIns+" group by agentname "
    SqlIns =SqlIns+"order by agentname "
    
    objrst1.Source =sqlins    
    objrst1.Open 
    
    if not (objrst1.EOF and objrst1.BOF) then
    
    %> 
   
   

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=bdate%>至<%=edate%>代理人销售有奖励客票明细表</title>
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
<p align="center"><b><font size="5" color="#000099"><font color=red><%=agentname%></font>销售<font color=red><%=company%></font>有奖励客票明细表</b>(承运日期：<font color=red><%=bdate%></font>至<font color=red><%=edate%></font>)</font></p>
<div align="center">
  <p align="left">  
  <center>
    <%
   
    'SqlIns ="exec cf_proc_temp '"& cstr(bdate) &"', '"& cstr(edate) &"'"
    
    %>  
 <table border="0" cellspacing="1" width="96%" height="1" bgcolor="#0000FF" bordercolor="#0000FF">
    <tr>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">代理人</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">航协号</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">航班日期</font></b></td>      
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">航班号</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">起飞城市</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">到达城市</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">票号</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">舱位</font></b></td>

      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">票价(净额)</font></b></td>            
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">销售日期</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">奖励费率1</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">奖励费1</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">奖励费率2</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">奖励费2</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">奖励费率3</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">奖励费3</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">奖励费率4</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">奖励费4</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">奖励费率5</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">奖励费5</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">奖励费合计</font></b></td>      
    </tr>

    
    
    <%


    objrst1.MoveFirst 
   while not objrst1.EOF  
    
    SqlIns =" select 代理人=agentname,航协号=agentsym,航班日期=flightdate,航班号=flightno, "
    SqlIns =SqlIns+"起飞城市=depcity,到达城市=arrcity, "
    SqlIns =SqlIns+"票号=ticketno,舱位=berthname,票价=price,销售日期=saledate, "    
    SqlIns =SqlIns+"奖励费率1=ar1,奖励额1=(ar1*price)/100,奖励费率2=ar2,奖励额2=(ar2*price)/100,"
    SqlIns =SqlIns+"奖励费率3=ar3,奖励额3=(ar3*price)/100,奖励费率4=ar4,奖励额4=(ar4*price)/100,奖励费率5=ar5,奖励额5=(ar5*price)/100, "
    SqlIns =SqlIns+"奖励费合计= ((ar1*price)+(ar2*price)+(ar3*price)+(ar4*price)+(ar5*price))/100  "
    SqlIns =SqlIns+" from "+ jdb
    SqlIns =SqlIns+" where "
    SqlIns =SqlIns+sqlwhere    
    SqlIns =SqlIns+"  and (ar1<>0 or ar2<>0 or ar3<>0 or ar4<>0 or ar5<>0 ) "    
    SqlIns =SqlIns+"  and agentname='" & trim(objrst1("代理人")) & "'  "    
    SqlIns =SqlIns+"order by agentname,agentsym,flightdate,depcity,arrcity,flightno,price desc "
    
    objrst.Source =sqlins
    
    objrst.Open 
    
'    if not (objrst.EOF and objrst.BOF) then

    
    objrst.MoveFirst 
    while not objrst.EOF 
    
   
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
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=objrst(8)%></font></td>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=objrst(9)%></font></td>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=objrst(10)%></font></td>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=objrst(11)%></font></td>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=objrst(12)%></font></td>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=objrst(13)%></font></td>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=objrst(14)%></font></td>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=objrst(15)%></font></td>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=objrst(16)%></font></td>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=objrst(17)%></font></td>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=objrst(18)%></font></td>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=objrst(19)%></font></td>
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=objrst(20)%></font></td>
    </tr>
    <%
    objrst.MoveNext 
    wend
    objrst.Close 
       


  %> 
    <tr>
      <td  height="1" bgcolor="burlywood"><font color="#0000FF"><%=trim(objrst1(0))%>合计</font></td>      
      <td  height="1" bgcolor="burlywood"><font color="#0000FF"></font></td>
      <td  height="1" bgcolor="burlywood"><font color="#0000FF"></font></td>
      <td  height="1" bgcolor="burlywood"><font color="#0000FF"></font></td>
      <td  height="1" bgcolor="burlywood"><font color="#0000FF"></font></td>
      <td  height="1" bgcolor="burlywood"><font color="#0000FF"></font></td>      
      <td  height="1" bgcolor="burlywood"><font color="#0000FF"></font></td>
  
      <td  height="1" bgcolor="burlywood"><font color="#0000FF"></font></td>
      <td  height="1" bgcolor="burlywood"><font color="#0000FF"></font></td>
      <td  height="1" bgcolor="burlywood"><font color="#0000FF"></font></td>
      <td  height="1" bgcolor="burlywood"><font color="#0000FF"></font></td>
      <td  height="1" bgcolor="burlywood"><font color="#0000FF"><%=objrst1(1)%></font></td>
      <td  height="1" bgcolor="burlywood"><font color="#0000FF"></font></td>
      <td  height="1" bgcolor="burlywood"><font color="#0000FF"><%=objrst1(2)%></font></td>
      <td  height="1" bgcolor="burlywood"><font color="#0000FF"></font></td>
      <td  height="1" bgcolor="burlywood"><font color="#0000FF"><%=objrst1(3)%></font></td>
      <td  height="1" bgcolor="burlywood"><font color="#0000FF"></font></td>
      <td  height="1" bgcolor="burlywood"><font color="#0000FF"><%=objrst1(4)%></font></td>
      <td  height="1" bgcolor="burlywood"><font color="#0000FF"></font></td>
      <td  height="1" bgcolor="burlywood"><font color="#0000FF"><%=objrst1(5)%></font></td>
      <td  height="1" bgcolor="burlywood"><font color="#0000FF"><%=objrst1("合计")%></font></td>
    </tr>
  
  
  <% 
  objrst1.movenext
  
  wend
  %>
  </table>
  
  <%
  
  objrst1.Close  
  
  else 
  	
  	%>
  <%=sqlins%>
  
  <p align="center"><b><font size="6" color="red">无符合条件的数据！</font></b></p>
  <p align="center">　</p>
<p align="center"><b><a href="./cwxs_index.asp"><font size="4" color="#0000CC">返 回</font></a></b></p>
  
  <%
  end if
      'objrst.Close  %>
  </center>
</div>


</body>
</HTML>
