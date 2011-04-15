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

sqldate1=" flidate>='"& bdate &"' and flidate<='"& edate &"' "


sqlwhere=sqldate+sqlco+sqlag+sqldep+sqlarr



'Response.Write sqlwhere
   

   
   
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=bdate%>至<%=edate%>代理人销售客票情况</title>
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
<p align="center"><b><font size="5" color="#000099"><font color=red> <%=agentname%></font>销售<font color=red><%=company%></font>客票情况</b>(承运日期：<font color=red><%=bdate%></font>至<font color=red><%=edate%></font>)</font></p>
<div align="center">
  <p align="left">  
  <center>
    <%
   
    'SqlIns ="exec cf_proc_temp '"& cstr(bdate) &"', '"& cstr(edate) &"'"
    
    SqlIns ="select 代理人=agent,起飞=khbz.depcity,到达=khbz.arrcity,"
    SqlIns =SqlIns+" 航班号=khbz.flightno,张数=zs,销售额=xse,班次=bc,客运班次=kybc,考核指标=khzb from "
    SqlIns =SqlIns+" (select * from "
    SqlIns =SqlIns+" (select * from "
    SqlIns =SqlIns+" (select agentname,depcity,arrcity,flightno,ZS=count(flightno),XSE=sum(price)"
    SqlIns =SqlIns+"  from ticketinfo "
    SqlIns =SqlIns+" where "
    SqlIns =SqlIns+sqlwhere
    SqlIns =SqlIns+" group by agentname,depcity,arrcity,flightno ) AS T "
    SqlIns =SqlIns+" left join "
    SqlIns =SqlIns+" (select flightno1=flightno,BC=sum(banci) from "
    SqlIns =SqlIns+" (select flightdate,flightno,banci=1 from ticketinfo  "
    SqlIns =SqlIns+" where " 
    SqlIns =SqlIns+sqldate
    SqlIns =SqlIns+" group by flightdate,flightno ) as B "
    SqlIns =SqlIns+" group by flightno "
    SqlIns =SqlIns+" ) as BC "
    SqlIns =SqlIns+" ON T.flightno=Bc.flightno1 ) as CXF "
    SqlIns =SqlIns+" left join "
    SqlIns =SqlIns+" ( "
    SqlIns =SqlIns+" select flightno2=flightno,KYbc=count(flightno) from "
    SqlIns =SqlIns+" (select distinct flidate,flightno from [10.254.0.46].[cftest].[dbo].keyunfromszx1 "
    SqlIns =SqlIns+" where "
    SqlIns =SqlIns+sqldate1
    SqlIns =SqlIns+ " ) as KY "
    SqlIns =SqlIns+" group by flightno  "
    SqlIns =SqlIns+" ) as KYBC  "
    SqlIns =SqlIns+" ON CXF.flightno='CZ'+KYBC.flightno2  "
    SqlIns =SqlIns+" ) as khb right join kh as khbz "
    SqlIns =SqlIns+" on khb.agentname=khbz.agent "
    SqlIns =SqlIns+" and khb.depcity=khbz.depcity "
    SqlIns =SqlIns+" and khb.arrcity=khbz.arrcity "
    SqlIns =SqlIns+" and khb.flightno=khbz.flightno "
    SqlIns =SqlIns+" order by agent,khbz.depcity,khbz.arrcity,khbz.flightno  "

    
    objrst.Source =sqlins
    
    objrst.Open 
    
    if not (objrst.EOF and objrst.BOF) then
    
    %>  
 <table border="0" cellspacing="1" width="96%" height="1" bgcolor="#0000FF" bordercolor="#0000FF">
    <tr>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">代理人</font></b></td>      
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">起飞航站</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">到达航站</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">航班号</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">客票数</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">销售金额</font></b></td>      
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">班次</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">客运班次</font></b></td>
      <td  height="1" bgcolor="#99CCFF"><b><font color="#990099">考核指标(张/班)</font></b></td>
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
      <td  height="1" bgcolor="#FFFFFF"><font color="#0000FF"><%=objrst(8)%></font></td>
    </tr>
    <%
    objrst.MoveNext 
    wend
    

  %> 

  </table>
  
  <% else %>
  
  <p align="center"><b><font size="6" color="red">无符合条件的数据！</font></b></p>
<p align="center">　</p>
<p align="center"><b><a href="./cwxs_index.asp"><font size="4" color="#0000CC">返 回</font></a></b></p>  
  
  <%
  end if
      objrst.Close  %>
  </center>
</div>


</body>
</HTML>
