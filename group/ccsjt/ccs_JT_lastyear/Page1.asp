<!--此特效来源来互联网,由 kudaa.com 收集整理-->
<!--可以折叠的树型菜单-->


<body bgcolor="cccccc">
<table width="200" height="250" border="0" align="center" cellpadding="0" cellspacing="0">


    <tr>
      <td align="center" valign="top" background="images/leftlist_bg.jpg"><script language=javascript id=clientEventHandlersJS>
<!--

 
function ShowFLT(i) 
{
    lbmc = eval('LM' + i);
    
    if (lbmc.style.display == 'none') 
    {
        //LMYC();
        
        lbmc.style.display = '';
    }
    else 
    {
        
        lbmc.style.display = 'none';
    }
}
//-->
      </script>
        <TABLE cellSpacing=0 cellPadding=0 width="88%" border=0>
        <TBODY>
                           <%
                           
                     Set objConn = Server.CreateObject("ADODB.Connection")
  objConn.Open Application("OledbStr") 
  Set objRst1=server.CreateObject ("ADODB.Recordset")
  objRst1.LockType=3
  objRst1.CursorType=3
  set objRst1.activeConnection=objConn
                  ' objrst1.Source ="select distinct kmshuom,kmcode from cwys_km where depar = '信息工程部' and kmshuom='其他业务支出-通讯费' order by kmcode,kmshuom" 
                  objrst1.Source ="select distinct kmshuom,kmcode from cwys_km where depar = '信息工程部' order by kmcode,kmshuom" 
    
                       'Response.Write(objrst1.Source)
                     
                       objrst1.Open 
                       j=1
                       while not objrst1.EOF
                  
                   %>
                       
          <TR>
       
          
            <TD style="PADDING-LEFT: 20px" background="" height=23><img src="images/a.gif" HEIGHT="23">
            <A onclick=javascript:ShowFLT(<%=j%>) 
                  href="javascript:void(null)"><%=trim(objrst1("kmshuom"))%></A> 
                  </TD>
                 
          </TR>
     
          <TR id=LM<%=j%> style="DISPLAY: none">
            <TD><TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
                <TBODY>
                 
               
                   <%
                     Set objConn = Server.CreateObject("ADODB.Connection")
  objConn.Open Application("OledbStr") 
  Set objRst=server.CreateObject ("ADODB.Recordset")
  objRst.LockType=3
  objRst.CursorType=3
  set objRst.activeConnection=objConn
                   objrst.Source ="select distinct kmshuom,kmcode,fkmshuom,fkmcode from cwys_km where depar = '信息工程部' and kmshuom='"&trim(objrst1("kmshuom"))&"' order by kmcode,kmshuom,fkmcode,fkmshuom" 
                       'Response.Write(objrst.Source)
                     
                       objrst.Open 
                       while not objrst.EOF
                   
                   %>

                  <TR>
                    <TD style="PADDING-LEFT: 40px" height=23> <A title=资料册 
                        href="ccs_input_index.asp?km=<%=trim(objrst("fkmshuom"))%>&km1=<%=trim(objrst("fkmcode"))%>&km2=<%=trim(objrst("kmshuom"))%>" 
                        ><font size="2"><%=trim(objrst("fkmshuom"))%></font></A> </TD>
                  </TR>
                  <TR>
                    <TD background="" height=3></TD>
                  </TR>
                                
           
                  
                    <%objrst.MoveNext %> 
                       <% wend %>    
                 
                    <%objrst.Close%>
                    
                  
                </TBODY>
            </TABLE></TD>
          </TR>
          
            <%
            objrst1.MoveNext 
            j=j+1
              %> 
                       <% wend %>    
                 
                    <%objrst1.Close%>
          
          
          
          
 
  </tbody>
</table>