<%@ Language=VBScript %>
<% 
  function Filter(value) 
      if value=-1 then
         filter="&nbsp"+"&nbsp"
      else
         filter=value
      end if      
  end function

   OledbStr_cf = "provider=sqloledb;server=10.254.0.46;database=cftest;uid=sa;pwd=;"  
   Set objConn_cf = Server.CreateObject("ADODB.Connection")
   objConn_cf.Open OledbStr_cf
   Set objRst=server.CreateObject ("ADODB.Recordset")
   objRst.LockType=3
   objRst.CursorType=3
   set objRst.activeConnection=objConn_cf    
   '连接数据库
   
   air=trim(Request.Form ("airline"))
   fli=trim(Request.Form ("flightno"))
   dep=trim(Request.Form ("depcity"))
   arr=trim(Request.Form ("arrcity"))
   dat=trim(Request.Form ("date"))
   if air="1" then
      airstr=""
   else
      airstr=" and airline='" & air & "'"
   end if
      
   if fli="全部" or fli="" then
      flistr=""
   else
      flistr=" and flight='" & fli & "'"
   end if   
   
   if dep="1" then
      depstr=""
   else
      depstr="depcity='" & dep & "'"
   end if   
      
   if arr="1" then
      arrstr=""
   else
      arrstr="arrcity='" & arr & "'"
   end if     
   
   if arrstr<>"" and depstr<>"" then
      depstr=depstr & " and "
   end if    
       
   
   objrst.Source= "select * from sale where date='"& dat &"' and " & depstr &  arrstr & airstr &  flistr  & " order by time "
   'objrst.Source ="select * from sale where" & airstr &  & flistr &  & depstr &  & arrstr & "order by time "
   
   
   'objrst.Source ="select * from sale where depcity='SZX' or arrcity='SZX' order by time"
   'objrst.Source ="select * from sale where depcity='SZX'order by time "
   'objrst.Source ="select * from sale where arrcity='SZX' order by time"
    'Response.Write objrst.Source 
   objrst.Open 
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

</HEAD>
<style type="text/css">

TD {
	FONT-FAMILY: 宋体; FONT-SIZE: 13px
}

TABLE {
	FONT-FAMILY: 宋体; FONT-SIZE: 13px
}

</style>
<body>
<% if not(objrst.BOF and objrst.EOF) then %>
<h1 align="center"><b><font face="长城行楷体" color="#6B2794">深圳进出港航班查询系统</font></b></h1>
<table border="1" cellspacing="0" width="756" cellpadding="0" bordercolor="#4b8ec5">
  <tr>
    <td align=center width="22" bgcolor="#abf8e2" ><STRONG><font color="#3333cc">时间</font></STRONG></td>
    <td align=center width="22" bgcolor="#abf8e2" ><STRONG><font color="#3333cc">航班号</font></STRONG></td>
    <td width="22" bgcolor="#abf8e2" ><STRONG><font color="#3333cc">起飞城市</font></STRONG></td>
    <td width="22" bgcolor="#abf8e2" ><STRONG><font color="#3333cc">到达城市</font></STRONG></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">机型</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >    
      <P align=center><STRONG><font color="#3333cc">头等舱</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">公务舱</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">普通舱</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">客座率</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">A</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">B</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">C</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">D</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">E</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">F</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">G</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">H</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">I</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">J</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">K</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">L</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">M</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">N</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">O</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">P</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">Q</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">R</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">S</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">T</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">U</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">V</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">W</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">X</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">Y</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">Z</font></STRONG></P></td>
  </tr>
  <%i=1 
    while not objrst.EOF %>
  <tr>
    <td ><%=objrst("time")%></td>
    <td ><%=trim(objrst("airline")+objrst("flight"))%></td>    
    <td><%=objrst("depcity")%></td>
    <td  ><%=objrst("arrcity")%></td>
    <td  ><%=objrst("flitype")%></td>
    <td align=middle><%=filter(objrst("fclass"))%></td>
    <td align=middle ><%=filter(objrst("bclass"))%></td>
    <td align=middle ><%=filter(objrst("eclass"))%></td>
    <td align=middle ><%=cstr(objrst("rate"))+"%"%></td>
    <td align=middle ><%=filter(objrst("A"))%></td>
    <td align=middle><%=filter(objrst("B"))%></td>
    <td align=middle ><%=filter(objrst("C"))%></td>
    <td align=middle><%=filter(objrst("D"))%></td>
    <td align=middle><%=filter(objrst("E"))%></td>
    <td align=middle><%=filter(objrst("F"))%></td>
    <td align=middle><%=filter(objrst("G"))%></td>
    <td align=middle><%=filter(objrst("H"))%></td>
    <td align=middle><%=filter(objrst("I"))%></td>
    <td align=middle><%=filter(objrst("J"))%></td>
    <td align=middle><%=filter(objrst("K"))%></td>
    <td align=middle><%=filter(objrst("L"))%></td>
    <td align=middle><%=filter(objrst("M"))%></td>
    <td align=middle><%=filter(objrst("N"))%></td>
    <td align=middle><%=filter(objrst("O"))%></td>
    <td align=middle><%=filter(objrst("P"))%></td>
    <td align=middle><%=filter(objrst("Q"))%></td>
    <td align=middle><%=filter(objrst("R"))%></td>
    <td align=middle><%=filter(objrst("S"))%></td>
    <td align=middle><%=filter(objrst("T"))%></td>
    <td align=middle><%=filter(objrst("U"))%></td>
    <td align=middle><%=filter(objrst("V"))%></td>
    <td align=middle><%=filter(objrst("W"))%></td>
    <td align=middle><%=filter(objrst("X"))%></td>
    <td align=middle><%=filter(objrst("Y"))%></td>
    <td align=middle><%=filter(objrst("Z"))%></td>
  </tr>
  <% objrst.MoveNext 
     i=i+1
     if (i mod 12 =0 ) then
    %>    
  <tr>
    <td align=center width="22" bgcolor="#abf8e2" ><STRONG><font color="#3333cc">时间</font></STRONG></td>
    <td align=center width="22" bgcolor="#abf8e2" ><STRONG><font color="#3333cc">航班号</font></STRONG></td>
    <td width="22" bgcolor="#abf8e2" ><STRONG><font color="#3333cc">起飞城市</font></STRONG></td>
    <td width="22" bgcolor="#abf8e2" ><STRONG><font color="#3333cc">到达城市</font></STRONG></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">机型</font></STRONG></P></td>    
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">头等舱</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">公务舱</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">普通舱</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">客座率</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">A</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">B</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">C</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">D</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">E</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">F</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">G</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">H</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">I</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">J</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">K</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">L</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">M</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">N</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">O</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">P</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">Q</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">R</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">S</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">T</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">U</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">V</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">W</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">X</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><STRONG><font color="#3333cc">Y</font></STRONG></P></td>
    <td width="23" bgcolor="#abf8e2" >
      <P align=center><font color="#3333cc"><STRONG>Z</STRONG></font></P></td>
  </tr>
    <%
     end if
     wend%>
</table>
<p align="center"><b><font color="#0000FF" size="4"><a href="queryindex.asp">返  回</a></font></b>
<% else %>
  <P>&nbsp</P>
<P>&nbsp</P>
<P>&nbsp</P>
  


<div align="center">

  <center>

  <table border="1" cellspacing="1" width="67%" bordercolor="#9933FF" height="87">
    <tr>
      <td width="100%" height="19">
        <p align="center"><b><font size="6" color="#FF33CC">没有您要查询的航班信息!</font></b></td>
    </tr>
    <tr>
      <td width="100%" height="56">
        <p align="center"><b><font color="#0000FF" size="5"><a href="queryindex.asp">返  回</a></font></b></td>
    </tr>
  </table>
  </center>
</div>

<% end if %>
</body>
</HTML>
