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
    FONT-FAMILY: 宋体; FONT-SIZE: 15px
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
  &nbsp;序号<input name="tid" readonly size ="10" value="<%=id%>">&nbsp;&nbsp; 
        姓名<input name="tname" size="10" value="<%=objrst("name")%>"> &nbsp;&nbsp;
        性别<select size="1" name="ssex">       
       <%if objrst("sex")=1 then%> 
    <option selected value="1">男</option>
    <option value="0">女</option>
       <%else%>
    <option value="1">男</option>
    <option selected value="0">女</option>       
       <%end if%>
  </select> &nbsp;&nbsp;
      出生日期
      <%if objrst("age")>0 then%>
      <input name="tberthday" size="15" value="<%=objrst("birthday")%>">
      <%else%>
      <input name="tberthday" size="15" value="" >
      <%end if%>
       &nbsp;&nbsp;
      血型<input name="Tbloodtype" size="4" value="<%=objrst("bloodtype")%>">    
 </p>   
  <p>&nbsp;办公电话<input name="Tofftel" size="21" value="<%=objrst("office_tel")%>" > 
  家庭电话<input name="Thometel" size="21" value="<%=objrst("home_tel")%>">   
  宿舍电话<input name="Tdormtel" size="19" value="<%=objrst("dorm_tel")%>">
  </p>  
  <p>&nbsp;手机<input name="Tmobile" size="16" value="<%=objrst("mobile_tel")%>"> 
  传呼<input name="Tbpcall" size="16" value="<%=objrst("bp_call")%>">       
  QQ号码<input name="Tqqcode" size="11" value="<%=objrst("qq_code")%>">&nbsp; 
  EMAIL<input name="Temail" size="21" value="<%=objrst("email")%>"></p>     
  <p>&nbsp;工作单位<input name="Tcorp" size="84" value="<%=objrst("corporation")%>" 
     ></p>
  <p>&nbsp;公司地址<input name="Toffaddr" size="84" value="<%=objrst("office_addr")%>"
     ></p>
  <p>&nbsp;家庭地址<input name="Thomeaddr" size="84" value="<%=objrst("home_addr")%>"
     ></p>
  <p>&nbsp;宿舍地址<input name="Tdormaddr" size="84" value="<%=objrst("dorm_addr")%>"
     ></p>
  <p>&nbsp;备注<input name="Tmemo" size="88" value="<%=objrst("memo")%>"
     ></p>
  <p align="center">&nbsp;所在城市<input name="Tcity" size="12" value="<%=objrst("city")%>">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  与本人关系
  <select size="1" name="srelation">   
  <% select case objrst("relation")
         case 0: Response.Write "<OPTION value=0 selected>朋友</OPTION>"
         case 1: Response.Write "<OPTION value=1 selected>小学同学</OPTION>"  
         case 2: Response.Write "<OPTION value=2 selected>初中同学</OPTION>"  
         case 3: Response.Write "<OPTION value=3 selected>高中同学</OPTION>"  
         case 4: Response.Write "<OPTION value=4 selected>大学同学</OPTION>"  
         case 5: Response.Write "<OPTION value=5 selected>同事</OPTION>"  
         case 6: Response.Write "<OPTION value=6 selected>亲属</OPTION>"    
     end select
  %>
    <option value=1>小学同学</option>
    <OPTION value=2>初中同学</OPTION>
    <OPTION value=3>高中同学</OPTION>
    <OPTION value=4>大学同学</OPTION>
    <OPTION value=0>朋友</OPTION>
    <OPTION value=5>同事</OPTION>
    <OPTION value=6>亲属</OPTION>
  </select>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
  关系密切指数
  <select size="1" name="Trelalevel">  
  <% select case objrst("relation_level")         
         case 1: Response.Write "<OPTION value=1 selected>极亲密</OPTION>"  
         case 2: Response.Write "<OPTION value=2 selected>较亲密</OPTION>"  
         case 3: Response.Write "<OPTION value=3 selected>亲密</OPTION>"  
         case 4: Response.Write "<OPTION value=4 selected>一般</OPTION>"           
     end select
  %>    
    <option value=1>极亲密</option>
    <OPTION value=2>较亲密</OPTION>
    <OPTION value=3>亲密</OPTION>
    <OPTION value=4>一般</OPTION>
  </select>  
  <br>
  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
  </p>
  <p>&nbsp;&nbsp;&nbsp; <input type="submit" value="提  交" name="B1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;  
  <input type="button" value="返  回" name="B2" LANGUAGE=javascript onclick="return B2_onclick()"></p>
</form>
</td>
  </tr>
</table>
  </center>
</div>
</body>


</html>
