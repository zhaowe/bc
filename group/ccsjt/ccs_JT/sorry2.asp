<%@ Language=VBScript %>
<%
   m_errorNo=session("errorNo")


   Set Conn1 = Server.CreateObject("ADODB.Connection")
   Conn1.Open Application("OledbStr") 
   
   
   Set Rs1=server.CreateObject ("ADODB.Recordset")
   Rs1.LockType=3
   Rs1.CursorType=3
   set Rs1.activeConnection=Conn1
      
   Rs1.Source="select * from error where Errornumber='" &trim(m_errorNo)& "'"
   Rs1.Open  


   if Rs1.eof and Rs1.bof then
      ErrorName="��Ǹ��ϵͳӵ����"
	  Solution="���Ժ��ٳ��ԣ�"
     else
      ErrorName=Rs1(1)
	  Solution=Rs1(2)
   end if  
                                                                                                   
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>�Բ������������ĳ�����Ϣ</title>
</HEAD>

<body >


<p align="center">
<table border="0" borderColor="#000000" borderColorDark="#c0c0c0" borderColorLight="#fdecec" cellPadding="0" cellSpacing="0" height="76" width="770">
  <tbody>
  <tr>
    <td height="90">
      <p align="center"><font color="#006600" size="5"><b><IMG border=0 height=60 src="../images/logo_air.gif" width=120>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b></font></p></td></tr></tbody></table>

<hr border=10>

   <table border='0' width='540'>
             <tr>
               <td width='440' align='left'><i>�Ϻ���˾��ҳ==&gt;</i></td>
               <td width='100' align='middle'><A class=ctrl_link href="javascript:history.back()"><font color='#408080'>����</font></a></td>
             </tr>
   </table>

   <center>
   <table border='6' width="80%" cellspacing='4' cellpadding='2' bordercolor='#99ccff'>
      <tr>
         <td bgcolor='#ffffff'>
               <table border=0 bgcolor='#ffffff'>
               </table>

                <p align="center"><font color="red" size=3>!<font> �Բ������������ĳ�����Ϣ��</p>

                <table align="center" border="1" width="100%">
                  <tr>
                    <td width="100" align="center"><font color="green">�������</font></td>
                    <td><font color="green"><%=m_errorNo%></font></td>
                  </tr>
                  <tr>
                    <td width="100" align="center"><font color="red">����ԭ��</font></td>
                    <td><font color="red" size=4><% =errorName%></font></td>
                  </tr>
                  <tr>
                    <td width="100" align="center"><font color="Brown">���;��</font></td>
                    <td><font color="Brown" size=4><%=Solution%></font></td>
                  </tr>
                </table>


                <table border=0 width="100%" bgcolor='#ffffff'>
                   <tr>
                      <td align='right'>
                          <A class=ctrl_link href="javascript:history.back()">
                          <font color='#408080'>����</font>
                          </a>
                          &nbsp;&nbsp;&nbsp;
                      </td>
                   </tr>
                   <tr>
                      <td align='middle'>
                          <hr color='#c0c0c0'>
                          <span class='tinyfont'>
                           �Ϻ����ڹ�˾.��Ȩ����
                           <br>
                           ��������κ�������飬������<A href="mailto:lizz@cs-air.com">����Ա
                           </a> 
                           <br>
                          </span>
                       </td>
                    </tr>
                </table>
       </tr>
   </table>
   
   </TD></TR></TABLE>
   </center>
  
  <br>
  <br>
  
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td>
      <p align=center>
      <font class="smallfont" color="#2167be" size="2">|
      <a href="http://web.cs-air.com">�Ϻ������ڲ���ҳ</a> |              
      <a href="http://www.cs-air.com">�Ϻ������ⲿ��ҳ</a> |              
      <a href="http://www.computerworld.com.cn">���������</a> |              

      <hr size="1" align="center">
      <span class="text">
      <font color="#000000" size="2">
      <p align="center">Copyright 2000 �й��Ϸ����գ����ţ����ڹ�˾<br>
        <font face="Arial, Helvetica, sans-serif"><strong>E-mail:</strong><A href="mailto:huzg@cs-air.com">huzg</A><A href="mailto:huzg@cs-air.com">@</A><A href="mailto:huzg@cs-air.com">cs-air.com</A></font></p>
           </font>
           </span></font>
      </td>
  </tr>
</table>

  
         
<%
  Conn1.Close
  set rs1 = nothing
  set Conn1 = nothing
%> 


<P></P>

</BODY>
</HTML>








