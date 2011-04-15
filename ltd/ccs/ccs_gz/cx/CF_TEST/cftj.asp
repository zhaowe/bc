<%@ Language=VBScript %>
<%

str=request.form("fen")
'str1=Request.QueryString("se")
response.write str
response.write "<BR>"
response.write "OK"


   OledbStr_cf = "provider=sqloledb;server=10.254.0.41;database=mastersystem;uid=sa;pwd=szx6275;"  
   Set objConn_cf = Server.CreateObject("ADODB.Connection")
   objConn_cf.Open OledbStr_cf
   Set objRst=server.CreateObject ("ADODB.Recordset")
   objRst.LockType=3
   objRst.CursorType=3
   set objRst.activeConnection=objConn_cf    
   '连接数据库
   

   objrst.Source ="update xjxscore set [end]=getdate(),score=" & str & " where no='196478'"  
   objrst.Open 

   response.redirect "kaoti20.asp"   

%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>me</title>
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
</body>
</HTML>
