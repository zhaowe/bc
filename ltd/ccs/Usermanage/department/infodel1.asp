

<html>


<head>
	<title>�Ŷӹ���ϵͳ</title>
	<link rel="stylesheet" type="text/css" href="style.css">
</head>

<body leftmargin=0 topmargin=0 marginheight="0" marginwidth="0" bgcolor="#afc0d0">



<!--��ӱ�//-->                                                            
  
 
  
  
      

<table border="0" cellspacing="0" cellpadding="0" width="100%" height="100%">
  <tr>
	<td valign="top">
      <table width="380" border="0" cellspacing="0" cellpadding="0">
      </table>


<table border="0" cellpadding="0" cellspacing="0" width="380">
<tr bgcolor="#AFC0D0" valign="top">
	<td width="380">
<p class="t01" style="font-size: 12px;"><font color="#FFFFFF"><img src="images/e05.gif" alt="" width="16" height="9" border="0"></font><b>ɾ����¼</b></p>

<!--ɾ������//-->                                                                          
<%
sub change() 
 
id=Request.QueryString("id")                                                                       
                                                                          
set conn_1=server.CreateObject("adodb.connection")                                                                            
    conn_1.Open Application("OledbStr")
 


%>


          
<%set rs_2=server.CreateObject("adodb.recordset")                                                                            
rs_2.CursorLocation=2                                                                            
sql5="delete FROM shouyichu_manage where record_id='"&id&"'"                                                                   
   
rs_2.Open sql5,conn_1,3,3,1 

%>
<font color="#000000" size="2"><STRONG><a href="infoin1.asp">�ü�¼��ɾ��,�뷵��</a></STRONG></font>



  <%     

End sub
%> 


  <%'������     
  id=Request.QueryString("id")                                                                                                                          
Select case Request.QueryString("todo")                                                                         
       case ""                                                                                                                                                                                                                   
                                                                                                                                              
         change() 
                                                             
End Select                                                                       
  %>                    
</td>
</tr>

</table>



  

	</td>
</tr>
</table>

</body>

