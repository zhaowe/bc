

<html>


<head>
	<title>部门管理</title>
	<link rel="stylesheet" type="text/css" href="style.css">
</head>

<body leftmargin=0 topmargin=0 marginheight="0" marginwidth="0" bgcolor="#afc0d0">



<!--添加表单//-->                                                            
  
 
  
  
      

<table border="0" cellspacing="0" cellpadding="0" width="100%" height="100%">
  <tr>
	<td valign="top">
      <table width="380" border="0" cellspacing="0" cellpadding="0">
      </table>


<table border="0" cellpadding="0" cellspacing="0" width="380">
<tr bgcolor="#AFC0D0" valign="top">
	<td width="380">
<p class="t01" style="font-size: 12px;"><font color="#FFFFFF"><img src="images/e05.gif" alt="" width="16" height="9" border="0"></font><b>删除记录</b></p>

<!--删除过程//-->                                                                          
<%
sub change() 
 
id=Request.QueryString("id")                                                                       
                                                                          
set conn_1=server.CreateObject("adodb.connection")                                                                            
    conn_1.Open Application("OledbStr")
 


%>


          
<%set rs_2=server.CreateObject("adodb.recordset")                                                                            
rs_2.CursorLocation=2                                                                            
sql5="delete FROM companylocale where companyid='"&id&"'"        
'Response.Write sql5
'Response.End                                                           
rs_2.Open sql5,conn_1,3,3,1

sql5="delete FROM companyinfo where companyid='"&id&"'"        
rs_2.Open sql5,conn_1,3,3,1 
sql5="delete FROM cwysglbm where companyid='"&id&"'"        
rs_2.Open sql5,conn_1,3,3,1


%>
<font color="#000000" size="2"><STRONG><a href="infoin1.asp">该记录已删除,请返回</a></STRONG></font>



  <%    

End sub
%> 


  <%'主过程    
  id=Request.QueryString("id")                                                                                  
                                                                                                                                 
  change() 
    %>                    
</td>
</tr>

</table>



  

	</td>
</tr>
</table>

</body>

