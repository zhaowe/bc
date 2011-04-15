<%@ Language=VBScript %>  
  <!-- #include virtual="sharecode/DataLink102.asp"-->
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

SumID=Request.form ("sumID")

edate=Request.Form ("edate")
bdate=dateadd("m",-1,dateadd("d",1,edate))

for i=1 to sumid-1 


yjdm="zkdm"+cstr(i)
vyjdm=Request.Form(yjdm)
yjdh="zkdh"+cstr(i)
vyjdh=Request.Form(yjdh)

   sql="insert into priceberthcode values( "
   sql=sql+" '" & bdate & "', '" & edate & "','所有',"
   sql=sql+" '" & vyjdm & "', '" & vyjdh & "')"
   
   objrst.Source =sql
   objrst.Open 
   
   'Response.Write sql 
   
next

%>
<P align="center">
数据提交成功！
</p>
</BODY>
</HTML>
