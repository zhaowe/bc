<%@ Language=VBScript %>  
  <!-- #include virtual="sharecode/DataLink102.asp"-->
<%



   ''OledbStr_cwxs = "provider=sqloledb;server=10.254.0.102;database=cwszx;uid=sa;pwd=123456;"  
   Set objConn_cf = Server.CreateObject("ADODB.Connection")
   objConn_cf.Open OledbStr_cwxs
   Set objRst=server.CreateObject ("ADODB.Recordset")
   objRst.LockType=3
   objRst.CursorType=3
   set objRst.activeConnection=objConn_cf    
   '�������ݿ�
   


%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft FrontPage 4.0">
<title>���������۲��񱨱�ϵͳ</title>
<style type="text/css">
.px12 {  font-size: 12px; line-height: 150%}
.px14 {  font-size: 14px; line-height: 150%}
.px16 {  font-size: 16px; line-height: 150%}
.px18 {  font-size: 18px; line-height: 150%}
.px24 {  font-size: 24px; line-height: 150%}
.px36 {  font-size: 36px; line-height: 150%}
.px48 {  font-size: 48px; line-height: 150%}
.px72 {  font-size: 72px; line-height: 150%}
body {  font-size: 12px; line-height: 150%}
p {  font-size: 12px; line-height: 150%}
td {  font-size: 9px; line-height: 150%}
input {  font-size: 12px; line-height: 150%}
select {  font-size: 12px; line-height: 150%}
.content4{FONT-SIZE:10PT; LINE-HEIGHT:9PT;}
.contentindex{font-family: "����";FONT-SIZE:9pt; LINE-HEIGHT:11pt;}
.enter {COLOR: #FFAF02; FONT-FAMILY: "����", "Arial", "Times New Roman"; FONT-SIZE: 11pt; TEXT-DECORATION: none ;font-weight: bold}
.head1{FONT-SIZE:11pt; LINE-HEIGHT:18pt; font-weight: bold; }
.head2{FONT-SIZE:10pt; LINE-HEIGHT:14pt; font-weight: bold; }
.contentsmall{FONT-SIZE:9pt; LINE-HEIGHT:12pt;}
.nav{FONT-SIZE:9pt; LINE-HEIGHT:10pt; color: #999999}
.content{FONT-SIZE:10pt; LINE-HEIGHT:14pt;color: #000000:#000000}
.news{FONT-SIZE:10pt; LINE-HEIGHT:14pt; color; color: #000000:#000000}
.contentbig{FONT-SIZE:11pt; LINE-HEIGHT:14pt;}
.info{  font-size: 9pt; line-height: 9pt;  color: #FFFFFF}
.footer{  font-size: 9pt; line-height: 12pt; font-weight: normal}
.search {  font-size: 10pt; line-height: 14pt; color: #ffffff; background-color: #75AEE3}
.whitehead {  font-size: 12pt; line-height: 15pt; color: #FFFFFF}
.whitecontent {  font-size: 10pt; line-height: 14pt; color: #ffffff}
.bgcolor {  background-color: #006797}
.leftline {  background-color: #FD7D04}
a:active {  color: #000000;; text-decoration: none}
a:visited {  color: #000000; font-weight: normal;; text-decoration: none}
a:link {  color: #000000; font-weight: normal; ; text-decoration: none}
a.homepage:link {  color: #000000; font-weight: normal;}
a.homepage:visited {  color: #000000; font-weight: normal;}
a.homepage:active {  color: #000000; font-weight: normal;}
a.homepage:hover {  color: #000000; font-weight: normal;}
</style>
</HEAD>
<body>


<div align="center" id="head" style="HEIGHT: 60px; LEFT: 3px; POSITION: absolute; TOP: 100px; WIDTH: 747px; Z-INDEX: 1">

<h1 align="center"><b><font color="#9900cc" face="������κ����" class="px36">���������۲��񱨱�ϵͳ</font></b></h1>

  <div align="center" style="width: 650; height: 94">
    <center>
    <div align="center">
    <form name="tijiao" method="post" action="RedirectRPT.asp">
      <table border="1" cellspacing="1" width="100%" bordercolor="#0099FF">
        <tr>
          <td ><input type="radio" value="1" checked  name="R1"><font class="px12">���������ۿ�Ʊ��ϸ(�н���);</font></td>
          <td ><input type="radio" value="2" name="R1"><font class="px12">���������ۿ�Ʊ����(�н���)</font></td>
          <td ><input type="radio" value="3" name="R1" ><font class="px12">���������ۿ�Ʊ���ܿ���</font></td>
          <td ><input type="radio" value="4" name="R1"><font class="px12">���������ۿ�Ʊ��ϸ(������)</font></td>
          <td ><input type="radio" value="5" name="R1"><font class="px12">�����˹�˾�����ѻ��ܱ�(�н���)</font></td>
          <td ><input type="radio" value="6" name="R1"><font class="px12">�����˹�˾�����ѻ��ܱ�(������)</font></td>
          <td ><input type="radio" value="7" name="R1"><font class="px12">�����˽����Ѹ���˾�ֺ��λ��ܱ�(�н���)</font></td>
          <td ><input type="radio" value="8" name="R1"><font class="px12">�����˷ֺ��ཱ����ϸ��(�н���)</font></td>          
          <td ><input type="radio" value="9" name="R1"><font class="px12">�����˽���������ܱ�(�н���)</font></td>          
          <td ><input type="radio" value="10"name="R1"><font class="px12">�����˽�������˾��̯��(�н���)</font></td>                    
          <td ><input type="radio" value="11"name="R1"><font class="px12">�����˽����Ѹ���˾��̯���ܱ�(�н���)</font></td> 
          <td ><input type="radio" value="12"name="R1"><font class="px12">�����˿�Ʊ��������ϸ��(�н���)</font></td>
          <td ><input type="radio" value="13"name="R1"><font class="px12">�����˺��ཱ������ϸ��(�н���)</font></td>
          <td ><input type="radio" value="14"name="R1"><font class="px12">������F/C��λ���ۻ��ܱ�(�н���)</font></td>    

        </tr>
        <tr>
          <td width="75%" colspan="14">
            <p align="center">&nbsp;<br>
                     <b><font color="#0000FF" class="px12">�������ڣ�</font></b>
            <select size="1" name="D1">
             <OPTION selected value="<%=cstr(year(date))%>"><%=cstr(year(date))%></OPTION>
             <OPTION  value="<%=cstr(year(date)-1)%>"><%=cstr(year(date)-1)%></OPTION>
             <OPTION  value="<%=cstr(year(date)-2)%>"><%=cstr(year(date)-2)%></OPTION>
             <OPTION  value="<%=cstr(year(date)-3)%>"><%=cstr(year(date)-3)%></OPTION>
            </select>
            <font class="px12">��</font>
            <select size="1" name="D2">
             <OPTION selected value="<%=cstr(month(date))%>"><%=cstr(month(date))%></OPTION>
             <% for i=1 to 12%>
             <OPTION  value="<%=i%>"><%=i%></OPTION>
             <% next %>            
            </select>
            <font class="px12">��</font>
            <select size="1" name="D3">
             <OPTION selected value="<%=cstr(day(date)-1)%>"><%=cstr(day(date)-1)%></OPTION>
             <% for i=1 to 31%>
             <OPTION  value="<%=i%>"><%=i%></OPTION>
             <% next %>               
            </select><font class="px12">��</font><font color="#0000FF"><b><font class="px12">��</font></b></font>     
            
            <select size="1" name="D4">
             <OPTION selected value="<%=cstr(year(date))%>"><%=cstr(year(date))%></OPTION>
             <OPTION  value="<%=cstr(year(date)-1)%>"><%=cstr(year(date)-1)%></OPTION>
             <OPTION  value="<%=cstr(year(date)-2)%>"><%=cstr(year(date)-2)%></OPTION>
            </select>
            <font class="px12">��</font>
            <select size="1" name="D5">
             <OPTION selected value="<%=cstr(month(date))%>"><%=cstr(month(date))%></OPTION>
             <% for i=1 to 12%>
             <OPTION  value="<%=i%>"><%=i%></OPTION>
             <% next %>            
            </select>
            <font class="px12">��</font>
            <select size="1" name="D6">
             <OPTION selected value="<%=cstr(day(date)-1)%>"><%=cstr(day(date)-1)%></OPTION>
             <% for i=1 to 31%>
             <OPTION  value="<%=i%>"><%=i%></OPTION>
             <% next %>               
            </select><font class="px12">��</font>                                    
         
            </p>
          </td>
        </tr>
        
        <tr>
          <td width="75%" colspan="14">
           <p align="center"> <br>
          <b><font color="#0000FF" class="px12">���˹�˾��</font></b>
         
         <%
         
            sql="select distinct arrcity from [123].dbo.AirlineData order by arrcity"
           objrst.Source =sql
           objrst.Open 
         %> 
          
          
            <select size="1" name="company">            
             <OPTION selected value="���й�˾">���й�˾</OPTION>
             <OPTION  value="�⹫˾">�⹫˾</OPTION>
             <% 
             objrst.MoveFirst 
             while not objrst.eof %>
             <OPTION  value="<%=objrst(0)%>"><%=objrst(0)%></OPTION>
             <% 
              objrst.MoveNext 
              wend              
              objrst.Close 
              %>               
            </select>          
          
          <b><font color="#0000FF" class="px12">&nbsp;�����ˣ�</font></b>
         
         <%
         
            sql="select distinct agentname from agentinfo order by agentname"
   objrst.Source =sql
   objrst.Open 
         %> 
          
          
            <select size="1" name="agent">            
             <OPTION selected value="���д�����">���д�����</OPTION>
             
             <% while not objrst.eof %>
             <OPTION  value="<%=objrst(0)%>"><%=objrst(0)%></OPTION>
             <% 
              objrst.MoveNext 
              wend              
              objrst.Close 
              %>               
            </select>
            
          <b><font color="#0000FF" class="px12">&nbsp;��ɺ�վ��</font></b>
         
         <%
         
            sql="select distinct arrcity from [123].dbo.AirlineData order by arrcity"
   objrst.Source =sql
   objrst.Open 
         %> 
          
          
            <select size="1" name="depcity">            
             <OPTION selected value="���к�վ">���к�վ</OPTION>
             <OPTION  value="SZX">SZX</OPTION>
             <% while not objrst.eof %>
             <OPTION  value="<%=objrst(0)%>"><%=objrst(0)%></OPTION>
             <% 
              objrst.MoveNext 
              wend              
              objrst.Close 
              %>               
            </select>  
            
            
            <b><font color="#0000FF" class="px12">&nbsp;���ﺽվ��</font></b>
         
         <%
         
            sql="select distinct arrcity from [123].dbo.AirlineData order by arrcity"
   objrst.Source =sql
   objrst.Open 
         %> 
          
          
            <select size="1" name="arrcity">            
             <OPTION selected value="���к�վ">���к�վ</OPTION>
             <OPTION  value="SZX">SZX</OPTION>
             <% while not objrst.eof %>
             <OPTION  value="<%=objrst(0)%>"><%=objrst(0)%></OPTION>
             <% 
              objrst.MoveNext 
              wend              
              objrst.Close 
              %>               
            </select>            
            
            </p>
          </td>
        </tr>        
        
        <tr>
          
          <td width="35%" colspan="14">
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
          <font color=red class="px14">������ʾ��ʽ</font>
          <br>  
          
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                      
            <input type="radio" value="1" checked name="R2"><font class="px12">���в������������β�в��ظ�������Ϣ</font>
            <br>
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
          
            <input type="radio" value="2" name="R2"><font class="px12">���в��������</font>
            <br>
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
          
            <input type="radio" value="3" name="R2"><font class="px12">β���ظ�������Ϣ</font>
            <br>
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
          
            <input type="radio" value="4" name="R2"><font class="px12">���в����������β���ظ�������Ϣ</font>
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
          
          </td>
          
          
        </tr>
        
        <tr>
          <td width="75%" colspan="14">
          

              <p align="center"><br><input type="submit" value="�� ��" name="B1"></p>
            
          </td>
        </tr>
        <tr>
          <td width="75%" colspan="14">
            <p align="center"><b><font color="#FF0000" class="px12">��ʾ:</font></b><font class="px12" color="#0000FF">������ƣ���Ϣ�������� �·� 6125��<br>&nbsp;&nbsp;&nbsp; 
            &nbsp;&nbsp;�����������µ�������ۿ� ������ 6192</font></td>
        </tr>
      </table>
      </form>
    </div>
    </center>
  </div>
  </center>
</div>

</body>

</HTML>
