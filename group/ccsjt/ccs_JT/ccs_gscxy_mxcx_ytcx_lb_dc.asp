<%@ Language=VBScript %>

<% 
 
 if trim(session("UID"))<>"" then
   dim objD
   set ObjD=server.CreateObject ("Com_UserManage1.ClsUserManage1")
       VerifyOk=objD.VerifyUserFunction (session("UID"),"CCS_GSCXY")
   if VerifyOk=false then
      session("errorNo")="000002"
      'Response.Redirect "../sorry/sorry.asp"
   end if   
 else
    session("errorNo")="000001"
    Response.Redirect "../sorry/sorry.asp"
 end if 
 
  'dep=Request.QueryString ("dep")




   Set objConn = Server.CreateObject("ADODB.Connection")
   objConn.Open Application("OledbStr") 
      
   Set objRst=server.CreateObject ("ADODB.Recordset")
   objRst.LockType=3
   objRst.CursorType=3
   set objRst.activeConnection=objConn
   
  Set objRst1=server.CreateObject ("ADODB.Recordset")
  objRst1.LockType=3
  objRst1.CursorType=3
  set objRst1.activeConnection=objConn
  
   Set objRst2=server.CreateObject ("ADODB.Recordset")
  objRst2.LockType=3
  objRst2.CursorType=3
  set objRst2.activeConnection=objConn

   



   emid=trim(session("emid"))
   loginid=trim(session("loginid"))

   sql="SELECT loginid,name,a.companyid,companyname FROM logininfo as a,companylocale as b "
   sql=sql+" where a.companyid=b.companyid and loginid='"& trim(session("loginid")) &"'"   
   objrst.open sql
   
   if objrst.eof and objrst.bof then
     Response.Write "¨oy?Y?à¨a3???¨a?ê?|ì???¨¨??¨°2?|ì?2???"
   else
     bm=objrst("companyname")
   end if
   
   objrst.Close 
  
  'bm="D???é1?è3¨?2?"
  
  'session("dep")=bm
  
  
  
  'year1=year(date)
  'date1=cdate(year1 & "-" & "1" & "-"&"1")
  'date2=cdate(year1 & "-" & "3" & "-"&"1")
  'date3=cdate(year1 & "-" & "6" & "-"&"1")
  'date4=cdate(year1 & "-" & "9" & "-"&"1")
  'date5=cdate(year1 & "-" & "12" & "-"&"1")
  'riqi=cdate(year1 & "-" & "1" & "-"&"1")
  'riqi1=cdate(year1 & "-" & "12" & "-"&"1")
  
  'km=Request.QueryString ("km")
  'session("km")=km



  lb=Request.QueryString ("month")
  
  syear=trim(Request.QueryString ("syear"))
  
  if syear="" then syear=year(date)


  
  syear=Request.Form("sel_year") 
  month1=Request.Form("sel_month_begin")
  month2=Request.Form("sel_month_end")  
  kmcode1=Request.Form("kmcode_1")  
  kmcode2=Request.Form("kmcode_2")  
  seldepart=Request.Form("selbm")  
  
  
  
  
  %>
<html>
<head>
<title>预算管理系统--全公司查询</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

<style type="text/css">
.px10 {  font-size: 10px; line-height: 150%}
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
.contentindex{font-family: "宋体";FONT-SIZE:9pt; LINE-HEIGHT:11pt;}
.enter {COLOR: #FFAF02; FONT-FAMILY: "宋体", "Arial", "Times New Roman"; FONT-SIZE: 11pt; TEXT-DECORATION: none ;font-weight: bold}
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
<script language="JavaScript">
<!--
function MM_preloadimagess() { //v3.0
  var d=document; if(d.imagess){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadimagess.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new images; d.MM_p[j++].src=a[i];}}
}

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapimages() { //v3.0
  var i,j=0,x,a=MM_swapimages.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}

function year1_onchange() {
  syear=year1.value ;
  surl='ccs_gscxy_main.asp?syear='+syear+'&month=';
  window.document.location.href(surl);
}

function selkm_onchange() {
  sbm=selbm.value ;
  syear=year1.value ;
  //if (sbm=="所有科目")
  //   surl='ccs_gscxy_main.asp?syear='+syear
  //else   
     surl='ccs_gscxy_index.asp?syear='+syear+'&km='+sbm;
  window.document.location.href(surl);
}

//-->
</script>

</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" text="#000000" link="#3333ff">
<%

   sql_lb=session("ytcx_sql")
   
   
   
   objrst.Source = sql_lb
   
   objRst.Open 
   if objrst.EOF and objrst.BOF   then %>       
     <p> 
     <p align="center"><font class="px12" color="black"><font color="crimson"><strong><%=syear%>年无科目预算数据
   <%else%>   




   
       <%
       
         '将表格数据输出成EXCEL
         Response.Buffer=true
         Response.ContentType = "application/vnd.ms-excel"
         Response.AddHeader "Content-Disposition", "attachment; filename="&year(now())&month(now())&day(now())&hour(now())&minute(now())&second(now())&".xls"


       
       %>     
            <table cellSpacing="1" cellPadding="0" width="100%">
              <tbody>
                
               <tr bgColor="#7dadc4" height="20">
                <td align="middle"><font class="px12" color="black">帐目时间</td>                                
                <td align="middle"><font class="px12" color="black">一级代码</td>                                                                         
                <td align="middle"><font class="px12" color="black">二级代码</td>                                
                <td align="middle"><font class="px12" color="black">费用科目</td>                                
                <td align="middle"><font class="px12" color="black">费用部门</td>                   
                <td align="middle"><font class="px12" color="black">核销日</td>                
                <td align="middle"><font class="px12" color="black">金额</td>
                <td align="middle"><font class="px12" color="black">费用说明</td>                
                <td align="middle"><font class="px12" color="black">录入人</td>
                </tr>
                <%
                i=1
                while not objrst.EOF 	
                 cbys="#ecf7fd"                 
                 if trim(objrst("isover"))="是" then cbys="#ef867a"
                 'if trim(objrst("isover"))="是" then cbys="red"                                
                %>
                <tr bgColor="<%=cbys%>" height="20">
                <td align="middle"><font class="px12" color="black"><%=year(objrst("mnytime"))%>年<%=month(objrst("mnytime"))%>月</td>
                <td align="middle"><font class="px12" color="black"><%=objrst("kmcode")%></td>                
                <td align="middle"><font class="px12" color="black"><%=objrst("fkmcode")%></td>
                <td align="middle"><font class="px12" color="black"><%=objrst("mnykm")%></td>                                                                              
                <td align="middle"><font class="px12" color="black"><%=objrst("mnydepm")%></td>                                                                                              
                <td align="middle"><font class="px12" color="black"><%=objrst("hxdate")%></td>                
                <td align="middle"><font class="px12" color="black"><%=objrst("price")%></td>                
                <td align="middle"><font class="px12" color="black"><%=objrst("mnynote")%></td>                
                <td align="middle"><font class="px12" color="black"><%=objrst("djname")%></td>                
                </tr>
                <%
                'i=i+1
                objrst.MoveNext 
                wend%> 
                
                
                
                <%objrst.Close %>
                
             
             </tbody></table></td></tr></tbody></table>
       
             
            

  <%end if%>    
 

 