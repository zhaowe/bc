

<html>

<script language="JavaScript">
function check()
{
if (confirm("你确定要提交吗？")==false)
  return false
}

</script>
<head>
	<title>部门管理</title>
	<link rel="stylesheet" type="text/css" href="style.css">
</head>

<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" bgcolor="#cccccc">
  <%dim people_number
people_number = session("emid")%>


<!--添加表单//-->                                                            
  
 
  
  

      
 <%sub add_data1()                                
 


 set conn_1=server.CreateObject("adodb.connection")                                                                          
    conn_1.Open Application("OledbStr")   
    set rs_2=Server.CreateObject ("ADODB.recordset")   
    rs_2.CursorLocation=2   
    'Sql_1="select max(record_ID) Id from shouyichu_use"   
    'rs_2.Open Sql_1,Conn_1,3,1,1    
    'id=rs_2("Id")+1
 %>
<table border="0" cellspacing="0" cellpadding="0" width="100%" height="100%" bgcolor="#666666">
  <tr>
	
	<td valign="top">
      <table width="780" border="0" cellspacing="0" cellpadding="0">
        
        <tr>
          <td height="29" background="images/right.jpg" algin="center">
            <table width="780" border="0" cellspacing="0" cellpadding="0">
              <tr>
               
              </tr>
            </table>
            <P align=center><STRONG><FONT size=1>部门维护</FONT></STRONG> 
        </P>
          </td>
        </tr>
      </table>


<table border="0" cellpadding="0" cellspacing="0" width="780">
<tr bgcolor="#e8e8e8" valign="top">
	<td width="780" agign="middle">

      
 <table border="0"             
            cellspacing="1" bgcolor="#ffffff" width="775" align=center height=20 
            style="FONT-SIZE: 12px" bordercolor="#ffffff"> 
   <form method="post" action="infoin.asp?todo=02" id=form1 name=form1 >
              <TBODY>


   
   <tr height=20 bgcolor="#e3e9ee" align="left">
   


      <td width=240>
          部门编号:&nbsp; <INPUT size=8 name=type1>
      </td>


      <td width=160>
          部门名称:<input name="text" size="8" 
                 >
      </td>

      <td align=left>
          <font color="#000080"><button id=button1 name=button1 type="submit" >增加</button></font>
      </td> 
   </tr>
        


</form>
   </tr>

</table>
        
<table border="0" cellspacing="0" cellpadding="0" width="780" align=center>
     <tr>

       <td align="left" bgcolor="#afc9e4">
        
           <table border="0" cellSpacing="1" cellPadding="0"      
                  width="402" style="FONT-SIZE: 12px" bgcolor=#ffffff align="left">
              <tr>                                                                          
                      <th nowrap width="20%" height="1" style="BACKGROUND-COLOR: #afc9e4" align="middle"><font color=black>部门编号</font></th>                                                                         
                      <th nowrap width="50%" height="1" style="BACKGROUND-COLOR: #afc9e4" align="middle"><font color=black>部门名称</font></th>  
                      <th nowrap style="BACKGROUND-COLOR: #afc9e4" height ="1" width="30%" align="middle"><font color=black>删除</font></th>  
              </tr>
           </table>
      </td>


      
   </tr>
</table>

<table border="0" cellspacing="0" cellpadding="0" width="780" align="center">
     <tr >

       <td align="left" bgcolor="#afc9e4">
        
           <table border="0" cellspacing="0" cellpadding="0" width="402" align="left">
                <tr align="left">
                  <td height="20" align="left" bgcolor="#afc9e4">
                   <iframe name="detail20" src="infoin1.asp" width=402 height=452 allowTransparency align=left frameborder="no">
                   </iframe>
                  </td>
                 </tr>
           </table>
        </td></tr>
          
        
</table></FORM>                     
</td>
</tr>
  <%
                                                                          
End sub                                                                          
  %>
</table><!--保存数据//-->
  <% 
                                                                        
Sub save_data()                                                                          


%>

<%
type1 = trim(Request.form("type1"))

text = trim(Request.form("text"))

if text="" or type1="" then%>
    <table border="0" cellspacing="0" cellpadding="0" width="100%" height="100%">
  <tr>
	<td width="50%" background="images/bg042.gif"><IMG height=1 alt="" src="images/px1.gif" width =1 border=0 ></td>
	<td valign="top">
      <table width="780" border="0" cellspacing="0" cellpadding="0">
        <tr>
         
        </tr>

        <tr>
          <td background="images/right.jpg">
            <table width="780" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="362"><center><IMG height=9 alt="" src="images/e06.gif" width =16 border=0 >&nbsp;&nbsp;<b><font size="2" color="#eceef2">欢迎您<%=session("loginname")%></font></b></center></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
  <center>
  
  <table>
    <tr align="middle">
      <td colspan="2" height="25" bgcolor="#cccccc" align="middle"><font color="blue" size="2"><%=type1%>内容为空，无法录入</font>
   </td>
    </tr>
    <font size="2">
    <tr>

      <td align="middle" bgcolor="#cccccc">
        <form action="infoin.asp" method="post" id="form2" name="form2">
          <input type="submit" value="重录" name="B5">
        </form>
      </td>
    </tr>
    </font>
  </table>
  </center>
<%else


  %>
    <%
set conn_1=server.CreateObject("adodb.connection")                                                                            
    conn_1.Open Application("OledbStr") 
set rs_8=server.CreateObject("adodb.recordset")                                                                            
                                                                          
rs_8.CursorLocation=2                                       

sql8="SELECT * FROM companylocale where companyid='"& type1 &"'"


   
rs_8.Open sql8,conn_1,3,3,1
if rs_8.EOF then 



  %>

   <%
set conn_1=server.CreateObject("adodb.connection")                                                                            
    conn_1.Open Application("OledbStr") 
set rs_1=server.CreateObject("adodb.recordset")                                                                            
                                                                          
rs_1.CursorLocation=2                                                                            
sql="SELECT * FROM companyinfo"                                                                            
rs_1.Open sql,conn_1,3,3,1
rs_1.AddNew
 
    Rs_1("companyid")=type1  
    Rs_1("Description")=text 
rs_1.Update
rs_1.Close

sql="SELECT * FROM companylocale"    
rs_1.Open sql,conn_1,3,3,1
rs_1.AddNew
    Rs_1("companyid")=type1  
    Rs_1("companyname")=text 
    Rs_1("locale")="zh"
rs_1.Update
rs_1.Close

sql="SELECT * FROM cwysglbm"    
rs_1.Open sql,conn_1,3,3,1
rs_1.AddNew
    Rs_1("companyid")=type1  
    Rs_1("companyname")=text 
    Rs_1("ysxs")="1"
rs_1.Update
rs_1.Close


  %>
    <table border="0" cellspacing="0" cellpadding="0" width="100%" height="100%">
  <tr>
	<td width="50%" background="images/bg042.gif"><IMG height=1 alt="" src="images/px1.gif" width =1 border=0 ></td>
	<td valign="top">
      <table width="780" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><IMG src="images/tuandui.jpg" width="100%" border=0></td>
        </tr>

        <tr>
          <td background="images/right.jpg">
            <table width="780" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="418"><A href="index.asp"><IMG height=29 src="images/but01.gif" width=76 border=0></A><A href="search.asp"><IMG height=29 src="images/but02.gif" width=76 border=0></A><A href="tuandui_manage.asp"><IMG height=29 src="images/but03.gif" width=76 border=0></A><A href="shouyichu_history.asp"><IMG height=29 src="images/but04.gif" width=76 border=0></A></td>
                <td width="362"><center><IMG height=9 alt="" src="images/e06.gif" width =16 border=0 >&nbsp;&nbsp;<b><font size="2" color="#eceef2">欢迎您<%=session("loginname")%></font></b></center></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
  <center>
  
  <table>
    <tr align="middle">
      <td colspan="2" height="25" bgcolor="#cccccc" align="middle"><font color="blue" size="2">提交成功 </font>
        <p>　</p></td>
    </tr>
    <font size="2">
    <tr>

      <td align="middle" bgcolor="#cccccc">
        <form action="infoin.asp" method="post" id="form2" name="form2">
          <input type="submit" value="继续提交" name="B5">
        </form>
      </td>
    </tr>
    </font>
  </table>
  </center>
  <%else
%>
    <table border="0" cellspacing="0" cellpadding="0" width="100%" height="100%">
  <tr>
	<td width="50%" background="images/bg042.gif"><IMG height=1 alt="" src="images/px1.gif" width =1 border=0 ></td>
	<td valign="top">
      <table width="780" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><IMG src="images/tuandui.jpg" width="100%" border=0></td>
        </tr>

        <tr>
          <td background="images/right.jpg">
            <table width="780" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="418"><A href="index.asp"><IMG height=29 src="images/but01.gif" width=76 border=0></A><A href="search.asp"><IMG height=29 src="images/but02.gif" width=76 border=0></A><A href="tuandui_manage.asp"><IMG height=29 src="images/but03.gif" width=76 border=0></A><A href="shouyichu_history.asp"><IMG height=29 src="images/but04.gif" width=76 border=0></A></td>
                <td width="362"><center><IMG height=9 alt="" src="images/e06.gif" width =16 border=0 >&nbsp;&nbsp;<b><font size="2" color="#eceef2">欢迎您<%=session("loginname")%></font></b></center></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
  <center>
  
  <table>
    <tr align="middle">

<% if type1="航线"   or type1="申请单位"   then %>

      <td colspan="2" height="25" bgcolor="#cccccc" align="middle"><font color="blue" size="2"><%=type1%>                                                                                                                                                      --<%=text%>已经存在</font>
<%  else %>
      <td colspan="2" height="25" bgcolor="#cccccc" align="middle"><font color="blue" size="2"><%=type1%>已经录入过，请先删除后再重新录入</font>

<%end if %>
   
   </td>
    </tr>
    <font size="2">
    <tr>

      <td align="middle" bgcolor="#cccccc">
        <form action="infoin.asp" method="post" id="form2" name="form2">
          <input type="submit" value="重录" name="B5">
        </form>
      </td>
    </tr>
    </font>
  </table>
  </center>
                        <CENTER>&nbsp;</CENTER>
  
<%

end if
end if
End sub
%>
  <%'主过程                                                                                                                          
Select case Request.QueryString("todo")                                                                         
       case ""                                                                                                                                                                                                                   
                                                                                                                                                
         add_data1()
         
       case"02" 
       save_data()  
                                                             
End Select                                                                       
  %>
  
<table border="0" cellpadding="0" cellspacing="0" width="780" height="10">
<tr>
	<td height="1"><IMG height=1 alt="" src="images/px1.gif" width =2 border=0 >&nbsp;</td>
</tr>
<center>

<tr bgcolor="#ee7b10">
	<td height="2" bgcolor="#000000" align="middle">
      <font face="宋体" size="2" color="#ebebeb">南方航空深圳分公司信息工程部技术室开发 v1.0 2006.4</font>                        
    </td>
</tr>
</table></CENTER>
	</td>
	<td width="50%" background="images/bg042.gif"><IMG height=1 alt="" src="images/px1.gif" width =1 border=0 ></td>
</tr>
</table></td></tr></table></td></tr></table></TD></TR></TBODY></TABLE>

</body>
