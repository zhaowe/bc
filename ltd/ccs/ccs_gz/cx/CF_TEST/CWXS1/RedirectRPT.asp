<%@ Language=VBScript %>
<%

   OledbStr_cf = "provider=sqloledb;server=10.254.0.102;database=cwszx;uid=sa;pwd=123456;"  
   Set objConn_cf = Server.CreateObject("ADODB.Connection")
   objConn_cf.Open OledbStr_cf
   Set objRst=server.CreateObject ("ADODB.Recordset")
   objRst.LockType=3
   objRst.CursorType=3
   set objRst.activeConnection=objConn_cf    
   '连接数据库
   

agentname=trim(Request.Form("agent"))
company=trim(Request.Form("company"))
depcity=trim(Request.Form("depcity"))
arrcity=trim(Request.Form("arrcity"))


typeid=Request.Form ("r1")

y=Request.Form ("d1")
m=Request.Form ("d2")
d=Request.Form ("d3")

y1=Request.Form ("d4")
m1=Request.Form ("d5")
d1=Request.Form ("d6")

selval=y+"-"+m+"-"+d
selval1=y1+"-"+m1+"-"+d1


selstr="bdate="+selval+"&"+"edate="+selval1+"&"+"ag="+agentname+"&"+"com="+company+"&"+"dep="+depcity+"&"+"arr="+arrcity


 sql="select top 1 flightdate from ticketinfo order by flightdate desc"
 objrst.Source =sql
 objrst.Open 
 
 flidate=objrst(0)
  
 objrst.Close
 set objrst=nothing
 
 
    if cdate(flidate)>=cdate(selval) then 
      select case typeid
         case 1:
             Response.Redirect "cwrptmx.asp?" & selstr
         case 2:
             Response.Redirect "cwrpt.asp?" & selstr
         case 3:
             Response.Redirect "cwrptkh.asp?" & selstr
         case 4:
             Response.Redirect "cwrptmx_all.asp?" & selstr
         case 5:
             Response.Redirect "cwrpt_company.asp?" & selstr
         case 6:
             Response.Redirect "cwrpt_company1.asp?" & selstr
         case 7:             
             Response.Redirect "cwrpt.asp?" & selstr
         case 8:
			 Response.Redirect "cwrpt.asp?" & selstr
      end select
    else

%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<body>

<p align="center">　</p>
<p align="center"><b><font color="#FF0000" size="5">选择日期超出报表数据范围!</font></b></p>
<p align="center">　</p>
<p align="center"><b><a href="./cwxs_index.asp"><font size="4" color="#0000CC">返 回</font></a></b></p>

</body>
</HTML>
<% end if%>