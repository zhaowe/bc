<%@ Language=VBScript %>
<%
   OledbStr_cf = "provider=sqloledb;server=10.254.0.46;database=cftest;uid=sa;pwd=szx6275;"  
   Set objConn_cf = Server.CreateObject("ADODB.Connection")
   objConn_cf.Open OledbStr_cf
   Set objRst=server.CreateObject ("ADODB.Recordset")
   objRst.LockType=3
   objRst.CursorType=3
   set objRst.activeConnection=objConn_cf    
   '连接数据库
   
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
  &nbsp;序号<input name="tid" readonly size ="10" value="<%=newid%>"> 姓名<input name="tname" size="10"     
     >    
  性别<select size="1" name="ssex">       
    <option selected value="1">男</option>
    <option value="0">女</option>
  </select> 出生日期<input name="tberthday" size="15"     
     > 血型<input name="Tbloodtype" size="4"     
     >    
  星座<input name="Txingzuo" size="4"     
     ></p>   
  <p>&nbsp;办公电话<input name="Tofftel" size="21" 
     > 家庭电话<input name="Thometel" size="21"    
     >   
  宿舍电话<input name="Tdormtel" size="19"    
     ></p>  
  <p>&nbsp;手机<input name="Tmobile" size="16" 
     > 传呼<input name="Tbpcall" size="16" >       
  QQ号码<input name="Tqqcode" size="11" >&nbsp; EMAIL<input name="Temail" size="21" ></p>     
  <p>&nbsp;工作单位<input name="Tcorp" size="84" 
     ></p>
  <p>&nbsp;公司地址<input name="Toffaddr" size="84" 
     ></p>
  <p>&nbsp;家庭地址<input name="Thomeaddr" size="84" 
     ></p>
  <p>&nbsp;宿舍地址<input name="Tdormaddr" size="84" 
     ></p>
  <p>&nbsp;现所在城市<input name="Tcity" size="18" 
     >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 与本人关系<select size="1" name="srelation">   
    <option value="1" 
        selected>小学同学</option><OPTION 
        value=2>初中同学</OPTION><OPTION value=3>高中同学</OPTION><OPTION 
        value=4>大学同学</OPTION><OPTION value=0>朋友</OPTION><OPTION value=5>同事</OPTION><OPTION value=6>亲属</OPTION>
  </select>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 关系密切指数<select size="1" name="Trelalevel">  
    <option value=1 
        selected>极亲密</option><OPTION 
        value=2>较亲密</OPTION><OPTION value=3>亲密</OPTION><OPTION 
        value=4>一般</OPTION>
  </select>  
  <br>
  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
  </p>
  <p>&nbsp;&nbsp;&nbsp; <input type="submit" value="提  交" name="B1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;  
  <input type="reset" value="重  写" name="B2"></p>
</form>
</td>
  </tr>
</table>
  </center>
</div>
</body>


</html>
