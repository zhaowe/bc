<!-- #include virtual="sharecode/DataLink102.asp"-->
<%
   Set objConn_cf = Server.CreateObject("ADODB.Connection")
   objConn_cf.Open OledbStr_cwxs
   Set objRst=server.CreateObject ("ADODB.Recordset")
   objRst.LockType=3
   objRst.CursorType=3
   set objRst.activeConnection=objConn_cf  
sqlcb="select count(*) from dbo.sysobjects where "
sqlcb=sqlcb+" id = object_id(N'[dbo].[tempresult]') "
sqlcb=sqlcb+" and OBJECTPROPERTY(id, N'IsUserTable') = 1"

objrst.Source =sqlcb
objrst.Open 
bexitID=objrst(0)
objrst.Close 


if bexitID>=1 then
sqlcb="drop table tempresult "
objrst.Source =sqlcb
objrst.Open 
end if
%>
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
   Set objRst0=server.CreateObject ("ADODB.Recordset")
   objRst0.LockType=3
   objRst0.CursorType=3
   set objRst0.activeConnection=objConn_cf    
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

'if cdate(edate)>=cdate("2006-01-01") and cdate(edate)<=cdate("2006-01-31") then
'   jdb=" ticketinfo_2006_01  "
'end if  

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
   

   
   
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=bdate%>至<%=edate%>代理人销售所有客票明细表</title>
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
<div align="center">
  <p align="left">  
  <center>
    <%
    'EXEC master..xp_cmdshell 'bcp "select * from [cftest].[dbo].keyunfromszx1" queryout d:\1222.txt -c -U sa -P szx6275'
    'SqlIns ="exec cf_proc_temp '"& cstr(bdate) &"', '"& cstr(edate) &"'"
    
    
    SqlIns =" select 代理人=agentname,航协号=agentsym,航班日期=flightdate,航班号=flightno, "
    SqlIns =SqlIns+"起飞城市=depcity,到达城市=arrcity, "
    SqlIns =SqlIns+"票号=ticketno,舱位=berthname,运价代号=pricecode,折扣代号=berthcode,票价=price,奖励费率1=ar1,奖励费率2=ar2,奖励费率3=ar3,奖励费率4=ar4,奖励费率5=ar5,奖励费率6=ar6,暗扣=ankou,销售日期=saledate into tempresult"
    SqlIns =SqlIns+" from "+ jdb
    SqlIns =SqlIns+" where "
    SqlIns =SqlIns+sqlwhere
    'SqlIns =SqlIns+"  and pricecode like 'yrt%' "    
    SqlIns =SqlIns+"order by agentname,agentsym,flightdate,depcity,arrcity,flightno,price desc "
    
    'Response.Write sqlins
    
    objrst0.Source = sqlins
    objrst0.Open 
    %>  
 <%
   Set objRst1=server.CreateObject ("ADODB.Recordset")
   objRst1.LockType=3
   objRst1.CursorType=3
   set objRst1.activeConnection=objConn_cf  
   SqlIns1 ="EXEC master..xp_cmdshell 'bcp [cwszx].[dbo].tempresult  out c:\Inetpub\ftproot\"+cstr(date)+".txt -c -U sa -P szx6275'"
   'SqlIns1 ="bcp cwszx..tempresult  out d:\Inetpub\ftproot\"+cstr(date)+".txt -c -U sa -P szx6275"
   objrst1.source=sqlins1
   'Response.Write sqlins1
   objrst1.open
   'objrst1.close
 %>
</center>
</div>
</p>
<center>	
<font color="#ff8000" size=5><b> 已将查询结果保存在ftp上。</b></font>
</center>
</body>
</HTML>
