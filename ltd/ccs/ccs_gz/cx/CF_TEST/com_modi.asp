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
   
   id=Request.QueryString("q")
   objrst.Source ="select * from cf_comm where id='"& id &"'"
   objrst.Open 

%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft FrontPage 4.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function B2_onclick() {
window.navigate("com_disp.asp")
}

//-->
</SCRIPT>
<style type="text/css">

TD {
    FONT-FAMILY: ����; FONT-SIZE: 15px
}
</style>
</head>
<body>
<div align="center">
  <center>
<table border="1" cellspacing="1" width="100%" bordercolor="#ffcc99">
  <tr>
    <td width="100%">

<form method="post" action="modify.asp" style="TEXT-ALIGN: center"  id=form1 name=form1>

  <p><br>
  &nbsp;���<input name="tid" readonly size ="10" value="<%=id%>">&nbsp;&nbsp; 
        ����<input name="tname" size="10" value="<%=objrst("name")%>"> &nbsp;&nbsp;
        �Ա�<select size="1" name="ssex">       
       <%if objrst("sex")=1 then%> 
    <option selected value="1">��</option>
    <option value="0">Ů</option>
       <%else%>
    <option value="1">��</option>
    <option selected value="0">Ů</option>       
       <%end if%>
  </select> &nbsp;&nbsp;
      ��������
      <%if objrst("age")>0 then%>
      <input name="tberthday" size="15" value="<%=objrst("birthday")%>">
      <%else%>
      <input name="tberthday" size="15" value="" >
      <%end if%>
       &nbsp;&nbsp;
      Ѫ��<input name="Tbloodtype" size="4" value="<%=objrst("bloodtype")%>">    
 </p>   
  <p>&nbsp;�칫�绰<input name="Tofftel" size="21" value="<%=objrst("office_tel")%>" > 
  ��ͥ�绰<input name="Thometel" size="21" value="<%=objrst("home_tel")%>">   
  ����绰<input name="Tdormtel" size="19" value="<%=objrst("dorm_tel")%>">
  </p>  
  <p>&nbsp;�ֻ�<input name="Tmobile" size="16" value="<%=objrst("mobile_tel")%>"> 
  ����<input name="Tbpcall" size="16" value="<%=objrst("bp_call")%>">       
  QQ����<input name="Tqqcode" size="11" value="<%=objrst("qq_code")%>">&nbsp; 
  EMAIL<input name="Temail" size="21" value="<%=objrst("email")%>"></p>     
  <p>&nbsp;������λ<input name="Tcorp" size="84" value="<%=objrst("corporation")%>" 
     ></p>
  <p>&nbsp;��˾��ַ<input name="Toffaddr" size="84" value="<%=objrst("office_addr")%>"
     ></p>
  <p>&nbsp;��ͥ��ַ<input name="Thomeaddr" size="84" value="<%=objrst("home_addr")%>"
     ></p>
  <p>&nbsp;�����ַ<input name="Tdormaddr" size="84" value="<%=objrst("dorm_addr")%>"
     ></p>
  <p>&nbsp;��ע<input name="Tmemo" size="88" value="<%=objrst("memo")%>"
     ></p>
  <p align="center">&nbsp;���ڳ���<input name="Tcity" size="12" value="<%=objrst("city")%>">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  �뱾�˹�ϵ
  <select size="1" name="srelation">   
  <% select case objrst("relation")
         case 0: Response.Write "<OPTION value=0 selected>����</OPTION>"
         case 1: Response.Write "<OPTION value=1 selected>Сѧͬѧ</OPTION>"  
         case 2: Response.Write "<OPTION value=2 selected>����ͬѧ</OPTION>"  
         case 3: Response.Write "<OPTION value=3 selected>����ͬѧ</OPTION>"  
         case 4: Response.Write "<OPTION value=4 selected>��ѧͬѧ</OPTION>"  
         case 5: Response.Write "<OPTION value=5 selected>ͬ��</OPTION>"  
         case 6: Response.Write "<OPTION value=6 selected>����</OPTION>"    
     end select
  %>
    <option value=1>Сѧͬѧ</option>
    <OPTION value=2>����ͬѧ</OPTION>
    <OPTION value=3>����ͬѧ</OPTION>
    <OPTION value=4>��ѧͬѧ</OPTION>
    <OPTION value=0>����</OPTION>
    <OPTION value=5>ͬ��</OPTION>
    <OPTION value=6>����</OPTION>
  </select>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
  ��ϵ����ָ��
  <select size="1" name="Trelalevel">  
  <% select case objrst("relation_level")         
         case 1: Response.Write "<OPTION value=1 selected>������</OPTION>"  
         case 2: Response.Write "<OPTION value=2 selected>������</OPTION>"  
         case 3: Response.Write "<OPTION value=3 selected>����</OPTION>"  
         case 4: Response.Write "<OPTION value=4 selected>һ��</OPTION>"           
     end select
  %>    
    <option value=1>������</option>
    <OPTION value=2>������</OPTION>
    <OPTION value=3>����</OPTION>
    <OPTION value=4>һ��</OPTION>
  </select>  
  <br>
  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
  </p>
  <p>&nbsp;&nbsp;&nbsp; <input type="submit" value="��  ��" name="B1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;  
  <input type="button" value="��  ��" name="B2" LANGUAGE=javascript onclick="return B2_onclick()"></p>
</form>
</td>
  </tr>
</table>
  </center>
</div>
</body>


</html>
