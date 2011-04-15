<%@ Language=VBScript %>
<% 


   
   air=trim(Request.Form ("airline"))
   fli=trim(Request.Form ("flightno"))
   dep=trim(Request.Form ("depcity"))
   arr=trim(Request.Form ("arrcity"))
   
   if air="1" then
      airstr=""
   else
      airstr=" and airline='" & air & "'"
   end if
      
   if fli="È«²¿" or fli="" then
      flistr=""
   else
      flistr=" and flightno='" & fli & "'"
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
   
   
   
   Response.Write "select * from sale where " & depstr &  arrstr & airstr &  flistr  & " order by time "
   'objrst.Source ="select * from sale where" & airstr &  & flistr &  & depstr &  & arrstr & "order by time "
   
   
   'objrst.Source ="select * from sale where depcity='SZX' or arrcity='SZX' order by time"
   'objrst.Source ="select * from sale where depcity='SZX'order by time "
   'objrst.Source ="select * from sale where arrcity='SZX' order by time"
   'objrst.Open 
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>

<P>&nbsp;</P>

</BODY>
</HTML>
