<%@ Language=VBScript %>
<%
   OledbStr_cf = "provider=sqloledb;server=10.254.0.46;database=cftest;uid=sa;pwd=szx6275;"  
   Set objConn_cf = Server.CreateObject("ADODB.Connection")
   objConn_cf.Open OledbStr_cf
   Set objRst=server.CreateObject ("ADODB.Recordset")
   objRst.LockType=3
   objRst.CursorType=3
   set objRst.activeConnection=objConn_cf    
   '�������ݿ�
   
   objrst.Source ="select * from cf_comm order by id desc"
   objrst.Open 
   if not(objrst.EOF and objrst.BOF) then
      maxid=objrst("id")
      newid=maxid+1
   else
      newid=1   
   end if   
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft FrontPage 4.0">
</head>
<body>
<div align="center">
  <center>
<table border="1" cellspacing="1" width="100%" bordercolor="#ffcc99">
  <tr>
    <td width="100%">

<form method="post" action="add.asp" style="TEXT-ALIGN: center" >

  <p><br>
  &nbsp;���<input name="tid" readonly size ="10" value="<%=newid%>"> ����<input name="tname" size="10"     
     >    
  �Ա�<select size="1" name="ssex">       
    <option selected value="1">��</option>
    <option value="0">Ů</option>
  </select> ��������<input name="tberthday" size="15"     
     > Ѫ��<input name="Tbloodtype" size="4"     
     >    
  ����<input name="Txingzuo" size="4"     
     ></p>   
  <p>&nbsp;�칫�绰<input name="Tofftel" size="21" 
     > ��ͥ�绰<input name="Thometel" size="21"    
     >   
  ����绰<input name="Tdormtel" size="19"    
     ></p>  
  <p>&nbsp;�ֻ�<input name="Tmobile" size="16" 
     > ����<input name="Tbpcall" size="16" >       
  QQ����<input name="Tqqcode" size="11" >&nbsp; EMAIL<input name="Temail" size="21" ></p>     
  <p>&nbsp;������λ<input name="Tcorp" size="84" 
     ></p>
  <p>&nbsp;��˾��ַ<input name="Toffaddr" size="84" 
     ></p>
  <p>&nbsp;��ͥ��ַ<input name="Thomeaddr" size="84" 
     ></p>
  <p>&nbsp;�����ַ<input name="Tdormaddr" size="84" 
     ></p>
  <p>&nbsp;�����ڳ���<input name="Tcity" size="18" 
     >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; �뱾�˹�ϵ<select size="1" name="srelation">   
    <option value="1" 
        selected>Сѧͬѧ</option><OPTION 
        value=2>����ͬѧ</OPTION><OPTION value=3>����ͬѧ</OPTION><OPTION 
        value=4>��ѧͬѧ</OPTION><OPTION value=0>����</OPTION><OPTION value=5>ͬ��</OPTION><OPTION value=6>����</OPTION>
  </select>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ��ϵ����ָ��<select size="1" name="Trelalevel">  
    <option value=1 
        selected>������</option><OPTION 
        value=2>������</OPTION><OPTION value=3>����</OPTION><OPTION 
        value=4>һ��</OPTION>
  </select>  
  <br>
  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
  </p>
  <p>&nbsp;&nbsp;&nbsp; <input type="submit" value="��  ��" name="B1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;  
  <input type="reset" value="��  д" name="B2"></p>
</form>
</td>
  </tr>
</table>
  </center>
</div>
</body>


</html>
