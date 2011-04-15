<%@ Language=VBScript %>
<% 
  '从数据库中取出datetime型的值, 将其转化为显示小时和分钟的字符串
 function gettime(dbdatetime)
dbdatetime=trim(dbdatetime)
if dbdatetime="" or isnull(dbdatetime) then
gettime=""
else
time_hour=hour(dbdatetime)
time_minute=minute(dbdatetime)
if time_hour<10 then time_hour="0"+cstr(time_hour) end if
if time_minute<10 then time_minute="0"+cstr(time_minute) end if
gettime=trim(cstr(time_hour)+":"+cstr(time_minute))
end if
end function
%>
<% 
 Set objConn = Server.CreateObject("ADODB.Connection")
   objConn.Open Application("OledbStr") 
   
   Set objRst=server.CreateObject ("ADODB.Recordset")
   objRst.LockType=3
   objRst.CursorType=3
   set objRst.activeConnection=objConn     
   mydate=dateadd("d",1,date)
   'objrst.Source = "select * from CDTABLE3 where date1='" & mydate & " ' order by gotime "
   'objRst.Open
   
   'Set objRst1=server.CreateObject ("ADODB.Recordset")
   'objRst1.LockType=3
   'objRst1.CursorType=3
   'set objRst1.activeConnection=objConn       
   objrst.Source = "select * from flighttomorrow where flightdate='" & mydate & " 'and changeflag<>3 and  fbflag=1 and begincity='圳' "
   objRst.Open
   
   if objrst.EOF and objrst.BOF   then
    objrst.Close  
     objrst.Source = "select * from flighttomorrow where flightdate='" & date & " ' "'order by gotime "
    objrst.Open 
    end if  
   'if objrst1.EOF and objrst1.BOF   then
   ' objrst1.Close  
    'objrst1.Source = "select * from flighttomorrow where flightdate='" & date & " 'and fbflag=1  "   
    'objrst1.Open 
    'end if          
         
         
         
   %> 
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>


<%
while not objrst.EOF
         
         if trim(objrst("fachetime"))="" or isnull(objrst("fachetime")) then
         qftime=gettime(objrst("qftime"))
         
         'qftime2=right(qftime,2)
         'qftime1=left(qftime,2)         
         
         'qftime=timeserial(qftime1,qftime2,0)        
         'qftime=timevalue(qftime)
          
         'if timevalue("0:0")=<qftime and qftime=<timevalue("10:00") then
          '   fctime=timeserial(qftime1-2,qftime2-10,0)
         'end if
         
         Response.Write "null"
         end if
         objrst.MoveNext 
         wend 

%>

<P>&nbsp;</P>

</BODY>
</HTML>
