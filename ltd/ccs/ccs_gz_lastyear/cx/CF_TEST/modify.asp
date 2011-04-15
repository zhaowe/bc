<%@ Language=VBScript %>
  <%
   OledbStr_cf = "provider=sqloledb;server=10.254.0.46;database=cftest;uid=sa;pwd=;"  
   Set objConn_cf = Server.CreateObject("ADODB.Connection")
   objConn_cf.Open OledbStr_cf
   Set objRst=server.CreateObject ("ADODB.Recordset")
   objRst.LockType=3
   objRst.CursorType=3
   set objRst.activeConnection=objConn_cf    
   '连接数据库
   
   
   id=cint(trim(Request.Form ("tid")))
   name=trim(Request.Form ("tname"))
   sex=trim(Request.Form ("ssex"))   
   birthday=trim(Request.Form ("tberthday"))
   if birthday="" then
      age=-1
   else
      age=year(date)-year(birthday)
   end if
   bloodtype=trim(Request.Form ("tbloodtype"))
   constellation=trim(Request.Form ("txingzuo"))
   office_tel=trim(Request.Form ("tofftel"))
   home_tel=trim(Request.Form ("thometel"))
   dorm_tel=trim(Request.Form ("tdormtel"))
   mobile_tel=trim(Request.Form ("tmobile"))
   BP_call=trim(Request.Form ("tbpcall"))
   email=trim(Request.Form ("temail"))
   QQ_code=trim(Request.Form ("tqqcode"))
   corporation=trim(Request.Form ("tcorp"))
   memo=trim(Request.Form ("Tmemo"))
   office_addr=trim(Request.Form ("Toffaddr"))
   home_addr=trim(Request.Form ("Thomeaddr"))
   dorm_addr=trim(Request.Form ("Tdormaddr"))
   city=trim(Request.Form ("Tcity"))
   relation=trim(Request.Form ("srelation"))
   relation_level=trim(Request.Form ("Trelalevel"))
   sql="update cf_comm set "
   sql=sql+"name= '" & name & "',sex='" & sex & "',age=" & age & ","
   sql=sql+"birthday='" & birthday & "',shengxiao='" & shengxiao & "',"
   sql=sql+"bloodtype='" & bloodtype & "',constellation='" & constellation & "',"
   sql=sql+"office_tel='" & office_tel & "',home_tel='" & home_tel & "',dorm_tel='" & dorm_tel & "',"
   sql=sql+"mobile_tel='" & mobile_tel & "',BP_call='" & BP_call & "',email='" & email & "',"
   sql=sql+"QQ_code='" & QQ_code & "',corporation='" & corporation & "',memo='" & memo & "',"
   sql=sql+"office_addr='" & office_addr & "',home_addr='" & home_addr & "',dorm_addr='" & dorm_addr & "',"
   sql=sql+"city='" & city & "',relation='" & relation & "',relation_level='" & relation_level & "' "
   sql=sql+"where id='" & id & "'"
   
   objrst.Source =sql
   objrst.Open 
   'sql="insert into cf_comm values( " & id & ", '" & name & "', '" & sex & "', " & age & ", '" & birthday & "', '" & bloodtype & "', '" & constellation & "', '" & office_tel & "', '" & home_tel & "', '" & dorm_tel & "', '" & mobile_tel & "', '" & BP_call & "', '" & email & "', '" & QQ_code & "', '" & corporation & "', '" & city & "', '" & office_addr & "', '" & home_addr & "', '" & dorm_addr & "', '" & relation & "', '" & relation_level & "')"
   'objrst.Open sql
   'Response.Write sql
   Response.Redirect "detail.asp?Q="+cstr(id)
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>

<P>&nbsp;</P>

</BODY>
</HTML>
