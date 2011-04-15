<%@ Language=VBScript %>


<%
   OledbStr_cf = "provider=sqloledb;server=10.254.0.102;database=123;uid=sa;pwd=123456;"  
   Set objConn_cf = Server.CreateObject("ADODB.Connection")
   objConn_cf.Open OledbStr_cf
 
   Set objRst=server.CreateObject ("ADODB.Recordset")
   objRst.LockType=3
   objRst.CursorType=3 
   set objRst.activeConnection=objConn_cf    
   
   Set objRst1=server.CreateObject ("ADODB.Recordset")
   objRst1.LockType=3
   objRst1.CursorType=3 
   set objRst1.activeConnection=objConn_cf     
   
   Set objRst2=server.CreateObject ("ADODB.Recordset")
   objRst2.LockType=3
   objRst2.CursorType=3 
   set objRst2.activeConnection=objConn_cf        
   
   'sql="select *  from hddata where len(hxdm)=11 order by flightno,flitype "
   'objrst.Source =sql
   'objrst.Open 
   
   sql1="select distinct flightno,flitype,hxdm,seatcount,banci  from hddata where len(hxdm)=11 order by flightno,flitype"
   objrst1.Source =sql1
   objrst1.Open 
   
   objrst1.MoveFirst 
   while not objrst1.EOF 
    
      fno=objrst1("flightno")
      ftype=objrst1("flitype")
      hxdm=objrst1("hxdm")
      seatcount=objrst1("seatcount")
      bc=objrst1("banci")
      
      rsAB=rs_AB(fno,ftype,hxdm)
      hdAB=hd_AB(fno,ftype,hxdm)
      rsAC=rs_AC(fno,ftype,hxdm)
      rsBC=rs_BC(fno,ftype,hxdm)
      hdBC=hd_BC(fno,ftype,hxdm)
      zws=seatcount*bc
      
      djrs=(rsab+rsac)*hdab+(rsac+rsbc)*hdbc
      
      djzws=zws*(hdab+hdbc)
      
      passrate=((rsab+rsac)*hdab+(rsac+rsbc)*hdbc)/(zws*(hdab+hdbc))
      
      sql_upd="update hxdata set  djrs=" & djrs & ",djzws=" & djzws & ",passrate=" & passrate & " where  flightno='" & fno & "' and flitype='" & ftype & "' and hxdm='" & hxdm & "' "
      objrst2.Source =sql_upd
      objrst2.Open 
     
      Response.Write passrate
      
      Response.Write "<br>"
      
      Response.Write "обр╩╦Ж!"
      
      Response.Write "<br>"
      
      objrst1.MoveNext 
      
     
   wend
   
   objrst1.Close
   
   
   
%>


<%


Function rs_AB( fno ,ftype ,hxdm )
  
  fno=trim(fno)
  ftype=trim(ftype)
  hxdm=trim(hxdm)
  fline=left(trim(objrst1("hxdm")),7)
  dep=left(fline,3)
  arr=right(fline,3)
      
  sql0="select personcount from hddata where len(hxdm)=11 and "
      
  sql2=" flightno='" & fno & "' and flitype='" & ftype & "' and depcity='" & dep & "' and arrcity='" & arr & "'"
      
  sql=sql0+sql2
      
  objrst.Source =sql
  objrst.Open 
    
  rs_AB=objrst(0)
  objrst.Close  

End Function 


function hd_AB(fno,ftype,hxdm)
  
  fno=trim(fno)
  ftype=trim(ftype)
  hxdm=trim(hxdm)
  fline=left(trim(objrst1("hxdm")),7)
  dep=left(fline,3)
  arr=right(fline,3)
      
  sql0="select distance from hddata where len(hxdm)=11 and "
      
  sql2=" flightno='" & fno & "' and flitype='" & ftype & "' and depcity='" & dep & "' and arrcity='" & arr & "'"
      
  sql=sql0+sql2
      
  objrst.Source =sql
  objrst.Open 
    
  hd_AB=objrst(0)
  objrst.Close  

end function 




function rs_AC(fno,ftype,hxdm)
  
  fno=trim(fno)
  ftype=trim(ftype)
  hxdm=trim(hxdm)
  fline=hxdm
  dep=left(fline,3)
  arr=right(fline,3)
      
  sql0="select personcount from hddata where len(hxdm)=11 and "
      
  sql2=" flightno='" & fno & "' and flitype='" & ftype & "' and depcity='" & dep & "' and arrcity='" & arr & "'"
      
  sql=sql0+sql2
      
  objrst.Source =sql
  objrst.Open 
    
  rs_AC=objrst(0)
  objrst.Close  

end function 


function rs_BC(fno,ftype,hxdm)
  
  fno=trim(fno)
  ftype=trim(ftype)
  hxdm=trim(hxdm)
  fline=right(trim(objrst1("hxdm")),7)
  dep=left(fline,3)
  arr=right(fline,3)
      
  sql0="select personcount from hddata where len(hxdm)=11 and "
      
  sql2=" flightno='" & fno & "' and flitype='" & ftype & "' and depcity='" & dep & "' and arrcity='" & arr & "'"
      
  sql=sql0+sql2
      
  objrst.Source =sql
  objrst.Open 
    
  rs_BC=objrst(0)
  objrst.Close  

end function 


function hd_BC(fno,ftype,hxdm)
  
  fno=trim(fno)
  ftype=trim(ftype)
  hxdm=trim(hxdm)
  fline=right(trim(objrst1("hxdm")),7)
  dep=left(fline,3)
  arr=right(fline,3)
      
  sql0="select distance from hddata where len(hxdm)=11 and "
      
  sql2=" flightno='" & fno & "' and flitype='" & ftype & "' and depcity='" & dep & "' and arrcity='" & arr & "'"
      
  sql=sql0+sql2
      
  objrst.Source =sql
  objrst.Open 
    
  hd_BC=objrst(0)
  objrst.Close  

end function 




%>
