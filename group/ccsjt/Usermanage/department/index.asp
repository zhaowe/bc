

<html>

<script language="JavaScript">
function check()
{
if (confirm("你确定要提交吗？")==false)
  return false
}

</script>
<head>
	<title>团队管理系统</title>
	<link rel="stylesheet" type="text/css" href="style.css">
</head>

<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" bgcolor="#cccccc">
  <%dim people_number
people_number = session("emid")%>


<!--添加表单//-->                                                            
  
 
  
  
      
 <%sub add_data1()                                
 
q = trim(Request.form("date1"))

unit1 = trim(Request.form("unit"))
journey1 = trim(Request.form("journey"))
flightway1 = trim(Request.form("flightway"))
startflight1 = trim(Request.form("startflight"))
returnflight1 = trim(Request.Form("returnflight")) 
 set conn_1=server.CreateObject("adodb.connection")                                                                          
    conn_1.Open Application("OledbStr")   
    set rs_2=Server.CreateObject ("ADODB.recordset")   
    rs_2.CursorLocation=2   
    Sql_1="select max(record_ID) Id from shouyichu_use"   
    rs_2.Open Sql_1,Conn_1,3,1,1    
    id=rs_2("Id")+1
 %>
<table border="0" cellspacing="0" cellpadding="0" width="100%" height="100%" bgcolor="#666666">
  <tr>
	<td width="50%" background="images/bg042.gif"><img src="images/px1.gif" width="1" height="1" alt border="0"></td>
	<td valign="top">
      <table width="780" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><img src="images/tuandui.jpg" border="0" width="100%"></td>
        </tr>

        <tr>
          <td height="29" background="images/right.jpg">
            <table width="780" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="418"><a href="index.asp"><img src="images/but011.gif" width="76" height="29" border="0"></a><a href="search.asp"><img src="images/but02.gif" width="76" height="29" border="0"></a><a href="tuandui_manage.asp"><img src="images/but03.gif" width="76" height="29" border="0"></a><a href="shouyichu_history.asp"><img src="images/but04.gif" width="76" height="29" border="0"></a><a href="infoin.asp"><img src="images/but05.gif" width="76" height="29" border="0"></a></td>
                <td width="362"><center><img src="images/e06.gif" width="16" height="9" alt border="0">&nbsp;&nbsp;<b><font size="2" color="#ECEEF2">欢迎您<%=session("loginname")%></font></b></center></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>


<table border="0" cellpadding="0" cellspacing="0" width="780">
<tr bgcolor="#e8e8e8" valign="top">
	<td width="780">

      
 <table border="0"  cellspacing="1" bgcolor="#FFFFFF" width="775" align=center height=20  style="font-size: 12px"  bordercolor="#ffffff" align=center > 
   <form method="POST" action="index.asp?todo=02" method="post" id=form1 name=form1>

   <tr height=20 bgcolor="#E3E9EE" align="left" >
                  <% 
       set conn_1=server.CreateObject("adodb.connection")                                                                          
    conn_1.Open Application("OledbStr")     
	set obs2=Server.CreateObject ("ADODB.recordset")   
    obs2.CursorLocation=2
    Sql_1="select * from shouyichu_manage where type1='申请单位'order by text"   
    obs2.Open Sql_1,Conn_1,3,1,1      
  %>
         <td width=125>
         
          &nbsp;单位 
          <select name="unit"> 
    <option value=<%=unit1%>><%=unit1%>   </option>      
   	                  <%                                                                         
	do while not obs2.EOF                                                                       
	%>
	<option value=<%=obs2("text")%>><%=obs2("text")%>   </option> 
	 <%                                                                                                                                            
	obs2.MoveNext                                                                      
	loop
  %>     
          </select> 
      </td>
       <%obs2.close%> 
       
        

      <td width=115>
          行程<input type="text" name="journey" size="10" value=<%=journey1%>>
      </td>
      
      
                  <% 
       set conn_1=server.CreateObject("adodb.connection")                                                                          
    conn_1.Open Application("OledbStr")     
	set obs2=Server.CreateObject ("ADODB.recordset")   
    obs2.CursorLocation=2
    Sql_1="select * from shouyichu_manage where type1='航线' order by text"   
    obs2.Open Sql_1,Conn_1,3,1,1      
  %>
         <td width=118>
         
          航线 
          <select name="flightway"> 
   <option value=<%=flightway1%>><%=flightway1%>   </option>       
   	                  <%                                                                         
	do while not obs2.EOF                                                                       
	%>
	<option value=<%=obs2("text")%>><%=obs2("text")%></option> 
	 <%                                                                                                                                            
	obs2.MoveNext                                                                      
	loop
  %>     
          </select> 
      </td>
       <%obs2.close%> 
      <td width=110>
          出发日<input name=date1 size="8" value="<%=q%>" readonly onClick="JavaScript:window.open('day.asp?form=form1&field=date1&oldDate='+this.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=260,height=200,top=230,left=400');">
      </td>       
      <td>
          航班<input type="text" name="startflight" size="4" value=<%=startflight1%>>
      </td>
      <td>         
          返回日<input name=date2 size="8" readonly onClick="JavaScript:window.open('day.asp?form=form1&field=date2&oldDate='+this.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=260,height=200,top=230,left=600');">
      </td>                      
      <td>
          航班<input type="text" name="returnflight" size="4" value=<%=returnflight1%>>
      </td>
   </tr>
   
   <tr height=20 bgcolor="#E3E9EE" align="left">
   
      <td>
          &nbsp;订座人数<input type="text" name="orderpeople" size="4">
      </td>
   
      <td>
          日期<input name=date3 size="8" value=<%=date()%> readonly onClick="JavaScript:window.open('day.asp?form=form1&field=date3&oldDate='+this.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=260,height=200,top=260,left=180');" size="16">
      </td>

      <td>
          时限<input name=ticketout size="8" value=<%=date()%> readonly onClick="JavaScript:window.open('day.asp?form=form1&field=ticketout&oldDate='+this.value,'','directorys=no,toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=260,height=200,top=260,left=180');" size="16">
      </td>

      <td>
          实出人数<input type="text" name="factpeople" size="3">
      </td>
   
      <td>
          编码<input type="text" name="code" size="4">
      </td>

      <td>
          状态
          <select name="station"> 
            <option value="等待">等待</option>
            <option value="取消">取消</option>
            <option value="已出票">已出票</option>
          </select>
      </td>


      <td >
          报价<input type="text" name="quote" size="3">
      </td>


   </tr>
        
   <tr height=20 bgcolor="#E3E9EE" align="left">
      <td>
          &nbsp;联系人<input type="text" name="linkman" size="5">
      </td>
      
      <td>
          方式<input type="text" name="linkway" size="10">
      </td>

      <td colspan=4>
          备注<input type="text" name="note" size="55">
      </td>

      <td align=center>
          <input type="submit" name="action" value="提交" onclick="return check()">
      </td> 
</form>
   </tr>

</table>
        

<table border="0" cellspacing="0" cellpadding="0" width="780" align=center>
     <tr >

       <td width="780" align="left" bgcolor="#afc9e4">
        

            </td>
      </tr>
</table>


<table border="0" cellspacing="0" cellpadding="0" width="780" align=center>
     <tr >
       <td width="780" height="20" align="left" bgcolor="#afc9e4">
        <iframe name="detail5" allowTransparency="true" src="indexdetail.asp" width=780 height=330 align=center frameborder="no">
        </iframe>
            </td>
      </tr>
</table>
<table border="0" cellspacing="0" cellpadding="0" width="780" align=center>
     <tr >
       <td width="780" height="20" align="left" bgcolor="#afc9e4">
        <iframe name="detail6" allowTransparency="true" src="shouyichu_detail.asp" width=780 height=95 align=center frameborder="no">
        </iframe>
            </td>
      </tr>
</table>


  </form>                     
</td>
</tr>
  <%
                                                                          
End sub                                                                          
  %>
</table>
<!--保存数据//-->
  <% 
                                                                        
Sub save_data()                                                                          


unit = trim(Request.form("unit"))
if unit=""then
unit ="无"
end if

journey = trim(Request.form("journey"))
if journey="" then
journey="无"
end if



startflight = trim(request.form("startflight"))
if startflight =""then
startflight = "无"
end if

returnflight = trim(request.form("returnflight"))
if returnflight=""then
returnflight = "无"
end if

ticketout= trim(request.form("ticketout"))
if  ticketout=""then
 ticketout="无"
end if

 orderpeople= trim(request.form("orderpeople"))
if  orderpeople=""then
orderpeople="0"
end if

 factpeople= trim(request.form("factpeople"))
if  factpeople=""then
 factpeople="0"
end if

 linkman= trim(request.form("linkman"))
if  linkman=""then
 linkman="无"
end if 

linkway= trim(request.form("linkway"))
if  linkway=""then
 linkway="无"
 end if 

quote= trim(request.form("quote"))
if  quote=""then
 quote="0"
end if

code= trim(request.form("code"))
if  code=""then
 code="无"
end if

note= trim(request.form("note"))
if  note=""then
 note="无"
end if

flightway= trim(request.form("flightway"))
if  flightway=""then
 flightway="无"
end if

station=trim(Request.form("station"))                                                                     
date1 = trim(Request.form("date1"))
date2 = trim(Request.form("date2"))
date3 = trim(Request.form("date3"))





  %>

  <%
set conn_1=server.CreateObject("adodb.connection")                                                                            
    conn_1.Open Application("OledbStr") 
set rs_1=server.CreateObject("adodb.recordset")                                                                            
                                                                          
rs_1.CursorLocation=2                                                                            
sql="SELECT * FROM shouyichu_use"                                                                            
                                                                           
                                                                       
rs_1.Open sql,conn_1,3,3,1
rs_1.AddNew
 
    Rs_1("unit")=unit  
    Rs_1("journey")=journey 
    Rs_1("flightway")=flightway 
    Rs_1("startday")=date1
    if date2="" then
    rs_1("returnday")=null
    else
    Rs_1("returnday")=date2
    end if 
    Rs_1("startflight")=startflight
    Rs_1("returnflight")=returnflight 
    Rs_1("orderpeople")=orderpeople 
    Rs_1("factpeople")=factpeople
    Rs_1("linkman")=linkman 
    Rs_1("linkway")=linkway 
    Rs_1("quote")=quote 
    Rs_1("code")=code 
    Rs_1("station")=station 
    Rs_1("orderday")=date3
    Rs_1("note")=note 
    Rs_1("ticketout")=ticketout
    Rs_1("flightway")=flightway                                                                       
rs_1("biaozhi")="录入"
 rs_1("lururen")=trim(session("emid"))
rs_1("lurutime")=date()



rs_1.Update


  %>
<%

End sub
%>
  <%'主过程                                                                                                                          
Select case Request.QueryString("todo")                                                                         
       case ""                                                                                                                                                                                                                   
                                                                                                                                                
         add_data1()
         
       case"02" 
       save_data()  
       add_data1()                                                      
End Select                                                                       
  %>
  
<table border="0" cellpadding="0" cellspacing="0" width="780" height="10">
<tr>
	<td height="1"><img src="images/px1.gif" width="1" height="1" alt border="0"></td>
</tr>
<center>
<tr bgcolor="#EE7B10">
	<td height="2" bgcolor="#000000" align="center">
      <font face="宋体" size="2" color="#EBEBEB">南方航空深圳分公司信息工程部技术室开发 v1.0 2006.4</font>                       
    </td>
</tr>
</table>
	</td>
	<td width="50%" background="images/bg042.gif"><img src="images/px1.gif" width="1" height="1" alt border="0"></td>
</tr>
</table>

</body>
