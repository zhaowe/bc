

<html>

<script language="JavaScript">
function check()
{
if (confirm("你确定要删除吗？")==false)
  return false
}

</script>
<head>
	<title>团队管理系统</title>
	<link rel="stylesheet" type="text/css" href="style.css">
</head>

<% '<body leftmargin=0 topmargin=0 marginheight="0" marginwidth="0" bgcolor="#ffffff"> %>

<body bgColor="transparent" leftmargin=0 topmargin=0 marginheight="0" marginwidth="0">


  <%dim people_number
people_number = session("emid")%>


<!--添加表单//-->                                                            
  
 
  
  
      


<%                                                                          
Sub sear_data()                                                                          

                                                          

%>


<%                     
set conn_1=server.CreateObject("adodb.connection")                                                                            
    conn_1.Open Application("OledbStr")      
set rs_2=server.CreateObject("adodb.recordset")                                                                            
rs_2.CursorLocation=2                                                                            
sql5="SELECT * FROM shouyichu_manage where type1='航线' "                                                                   
sql5=sql5&" order by record_id desc"    
rs_2.Open sql5,conn_1,3,3,1 

%>




 <%if rs_2.EOF then %>                                                                                                   


<table border="0" cellSpacing="1" cellPadding="0"  width="202" style="font-size: 12px" bgcolor=#ffffff >

   <tr>                                                                    
                                                                      
      <td align=center nowrap width="60" bgcolor="#ebebeb" height="20">&nbsp;</td>                                                                        
      <td align=center nowrap width="70" bgcolor="#ebebeb" height="20">&nbsp;</td>
      <td align=center nowrap width="40" bgcolor="#ebebeb" height="20">&nbsp;</td>
      
   </tr> 

</table>                                                                          
<% else %>

<%                                                                        
do while not rs_2.EOF                                                                        
%>

<table border="0" cellSpacing="1" cellPadding="0"  width="202" style="font-size: 12px" bgcolor=#ffffff >

<tr valign="top">

      <td align=center nowrap width="60" bgcolor="#ebebeb" height="20"><%=rs_2("type1")%></td>                                                                        

      <td align=center nowrap width="70" bgcolor="#ebebeb" height="20"><%=rs_2("text")%></td>

  <td align=center width="40" bgcolor="#ebebeb" height="20"><a href="infodel.asp?id=<%=rs_2("record_id")%> " onclick="return check()" target="detail21">删除</a>   
      
</tr> 
                                                                    
<%                                                                                                                                         
rs_2.MoveNext                                                                      
loop%>
</tr> 
</table> 

<%end if%> 



<%rs_2.close
 set Rs_2=nothing   %>  

  
   
                                                                

<%end sub%>
<%call sear_data()%>              

</body>

