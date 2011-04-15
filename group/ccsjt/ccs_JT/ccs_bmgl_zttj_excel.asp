<%@ Language=VBScript %>
<%


km=Request.QueryString ("km")
%>
<% 
 
if trim(session("UID"))<>"" then
  dim objD
  set ObjD=server.CreateObject ("Com_UserManage1.ClsUserManage1")
       VerifyOk_gsld=objD.VerifyUserFunction (session("UID"),"ccs_ldcx")
    'if VerifyOk_gsld=false then
      'session("errorNo")="000002"
      'Response.Redirect "../../sorry/sorry.asp"
    'else
    '   bm=Request.QueryString ("bm") 
    'end if   
 else
   session("errorNo")="000001"
   Response.Redirect "../../sorry/sorry.asp"
end if 
 
  'dep=Request.QueryString ("dep")

set conn_1=server.CreateObject("adodb.connection")                                                                            
    conn_1.Open Application("OledbStr") 
e=trim(session("emid"))
t=trim(session("loginid"))
'Response.Write(e)
set conn_1=server.CreateObject("adodb.connection")                                                                            
    conn_1.Open Application("OledbStr")                                                                           

set rs_1=server.CreateObject("adodb.recordset")                                                                            
rs_1.CursorLocation=2                                                                            
sql1="SELECT * FROM logininfo WHERE loginid='"& t &"'"                                                                            
rs_1.Open sql1,conn_1,3,3,1
if not rs_1.EOF then
cid=rs_1("companyid")
name=rs_1("name")
else
Response.Write("请重新登陆")
end if

set rs_3=server.CreateObject("adodb.recordset")
rs_3.CursorLocation=2
sql3="select * from companylocale where companyid='"&cid&"'"
rs_3.Open sql3,conn_1,3,3,1

f=trim(rs_3("companyname"))


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
     Response.Write "数据表出错，登录人找不到部门"
   else
     depart=trim(objrst("companyname"))
   end if
   objrst.Close 
   
   bm=trim(Request.QueryString ("bm"))      
    
    if VerifyOk_gsld=true and bm<>"" then
       session("dep")=bm
    else
       bm=depart
       session("dep")=depart
    end if
   
   'if bm="公司领导" then 
    '  bm=Request.QueryString ("bm")
   'else
   '   session("dep")=bm
   'end if      
  
   'bm="信息工程部"
   
    bm=session("dep")
   
   
  
  
  
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
  'Response.Write lb
  syear=trim(Request.QueryString ("syear"))
  
  if syear="" then syear=year(date)
  
  %>
<html>
<head>
<title>公司预算管理系统--<%=bm%></title>
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
  //sbm=selbm.value ;
  sbm='<%=bm%>';
  syear=year1.value ;
  surl='ccs_bmgl_zttj.asp?syear='+syear+'&bm='+sbm;
  window.document.location.href(surl);
}



function selbm_onchange() {
  sbm=selbm.value ;
  syear=year1.value ;
  if (sbm=="所有部门")
     surl='ccs_gsldcx_main.asp?syear='+syear
  else   
     surl='ccs_bmgl_zttj.asp?syear='+syear+'&bm='+sbm;
  window.document.location.href(surl);
}

function ShowFLT(i) 
{
    lbmc = eval('LM' + i);
    
    if (lbmc.style.display == 'none') 
    {
        //LMYC();
        
        lbmc.style.display = '';
    }
    else 
    {
        
        lbmc.style.display = 'none';
    }
}
//-->
</script>

</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" text="#000000" link="#3333FF" >
<%km=Request.QueryString ("km")
km2=Request.QueryString ("km2")
fkmcode1=Request.QueryString ("km1")%>
               
          
                 
                                
            

  
  
  <%
   sql_ze="ze=a.niandu,"
   sql_fy=""
   cxsj=cstr(syear)+"全年"
   
   select case  cstr(lb)
     case "1":
           sql_ze="ze=a.jan,"
           sql_fy=" and month(mnytime)=1 "
           cxsj="一月份"
     case "2":
           sql_ze="ze=a.feb,"
           sql_fy=" and month(mnytime)=2 "
           cxsj="二月份"
     case "3":
           sql_ze="ze=a.mar,"
           sql_fy=" and month(mnytime)=3 "
           cxsj="三月份"
     case "4":
           sql_ze="ze=a.apr,"
           sql_fy=" and month(mnytime)=4 "
           cxsj="四月份"
     case "5":
           sql_ze="ze=a.may,"
           sql_fy=" and month(mnytime)=5 "
           cxsj="五月份"
     case "6":
           sql_ze="ze=a.jun,"
           sql_fy=" and month(mnytime)=6 "
           cxsj="六月份"
     case "7":
           sql_ze="ze=a.jul,"
           sql_fy=" and month(mnytime)=7 "
           cxsj="七月份"
     case "8":
           sql_ze="ze=a.aug,"
           sql_fy=" and month(mnytime)=8 "
           cxsj="八月份"
     case "9":
           sql_ze="ze=a.sep,"
           sql_fy=" and month(mnytime)=9 "
           cxsj="九月份"
     case "10":
           sql_ze="ze=a.oct,"
           sql_fy=" and month(mnytime)=10 "
           cxsj="十月份"
     case "11":
           sql_ze="ze=a.nov,"
           sql_fy=" and month(mnytime)=11 "
           cxsj="十一月份"
     case "12":
           sql_ze="ze=a.dece,"
           sql_fy=" and month(mnytime)=12 "
           cxsj="十二月份"
           
     case "jd1":
           sql_ze="ze=a.jan+a.feb+a.mar,"
           sql_fy=" and (month(mnytime)>=1 and month(mnytime)<=3 ) "
           cxsj="截止一季度"
     case "jd2":
           sql_ze="ze=a.jan+a.feb+a.mar+a.apr+a.may+a.jun,"
           sql_fy=" and (month(mnytime)>=1 and month(mnytime)<=6 ) "
           cxsj="截止二季度"
     case "jd3":
           sql_ze="ze=a.jan+a.feb+a.mar+a.apr+a.may+a.jun+a.jul+a.aug+a.sep,"
           sql_fy=" and (month(mnytime)>=1 and month(mnytime)<=9 ) "
           cxsj="截止三季度"
     case "jd4":
           sql_ze="ze=a.jan+a.feb+a.mar+a.apr+a.may+a.jun+a.jul+a.aug+a.sep+a.oct+a.nov+a.dece,"
           sql_fy=" and (month(mnytime)>=1 and month(mnytime)<=12 ) "
           cxsj="截止四季度"
     case "全年":
           sql_ze="ze=a.niandu,"
           sql_fy=""
           cxsj=cstr(syear)+"全年"      
     
   end select 
   
  
  %>
      

            <%
       
         '将表格数据输出成EXCEL
         Response.Buffer=true
         Response.ContentType = "application/vnd.ms-excel"
         Response.AddHeader "Content-Disposition", "attachment; filename="&year(now())&month(now())&day(now())&hour(now())&minute(now())&second(now())&".xls"


       
       %>  
      
      

   
   
   <%
   
  
   
   
   
   
   
   sql="select ys_year,fkmshuom,ze,je=sum(je),bl=sum(je)/ze from " 
   sql=sql+"( " 
   sql=sql+"select a.*,je=isnull(b.price,0) from  " 
   sql=sql+"( " 
   sql=sql+"select a.depar,a.ys_year,a.fkmcode,"
   'sql=sql+"a.niandu,"
   sql=sql+sql_ze
   sql=sql+"b.fkmshuom from cwys_ed as a,cwys_km as b  " 
   sql=sql+"where a.depar='"& bm &"' and a.ys_year='"& syear &"' and a.fkmcode=b.fkmcode and a.ys_year=b.nian and a.depar=b.depar " 
   sql=sql+") as a " 
   sql=sql+"left join cwys_infoin b on a.fkmcode=b.mnykmcode and a.depar=b.mnydepm and a.ys_year=b.mnyyear and ifhandin='是' "
   sql=sql+sql_fy 
   sql=sql+") as c " 
   sql=sql+"group by ys_year,fkmshuom,fkmshuom,ze " 

   'Response.Write sql
   
   objrst.Source = sql
   'Response.Write sql
   objRst.Open 
   if objrst.EOF and objrst.BOF   then %>    
     <P> 
     <P align=center><font class="px12" color="black"><FONT color=crimson><STRONG><%=syear%>年无科目预算数据
   <%else%>
   
<table align="center"  cellSpacing="0" cellPadding="0" width="777" border="0">
  <tbody>
  <tr>
    <td colSpan="2" height="3"></td></tr>
  <tr>
    <td vAlign="top" width="73%">
      <table style="BORDER-RIGHT: #4983a0 1px solid; BORDER-TOP: #4983a0 1px solid; BORDER-LEFT: #4983a0 1px solid; BORDER-BOTTOM: #4983a0 1px solid" height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
        <tbody>
  
        <tr>
          <td vAlign="top">
            <table cellSpacing="1" cellPadding="0" width="100%">
              <tbody>
               <tr bgColor="#7dadc4" height="20">
                <td align="middle"  ><font class="px12" color="black">帐目时间</td>                                
                <td align="middle"  ><font class="px12" color="black">费用科目</td>                                
                <td align="middle"  ><font class="px12" color="black">指标</td>
                <td align="middle"  ><font class="px12" color="black">完成情况</td>
                <td align="middle"  ><font class="px12" color="black">剩余情况</td>
                <td align="middle"  ><font class="px12" color="black">完成率</td>                
                </tr>
                <%
                
                while not objrst.EOF 
                 cbys="#ecf7fd"
                 if objrst("bl")>1 then cbys="#ef867a"
                 
                %> 
                <tr bgColor="<%=cbys%>" height="20">                
                <td align="middle" ><font class="px12" color="black"><%=cxsj%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("fkmshuom")%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("ze")%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("je")%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("ze")-objrst("je")%></td>
                <td align="middle" ><font class="px12" color="black"><%=formatpercent(objrst("bl"),2,-1)%></td>                                                                                                                                
                
                </tr>
                <%
                objrst.MoveNext 
                wend%> 
                <%objrst.Close %>
             
     <%
   sql_hj="select ze=sum(ze),je=isnull(sum(je),0),bl=isnull(sum(je)/sum(ze),0) from ( "  
   sql_hj=sql_hj+sql
   sql_hj=sql_hj+" ) as sy "
   
   objrst.Source =sql_hj
   objrst.Open    
     %>        
               
                <tr bgColor="#ecf7fd" height="20">
                <td align="middle" ><font class="px12" color="blue"><%=cxsj%>合计</td>
                <td align="middle" ><font class="px12" color="blue">所有科目</td>
                <td align="middle" ><font class="px12" color="blue"><%=objrst("ze")%></td>
                <td align="middle" ><font class="px12" color="blue"><%=objrst("je")%></td>
                <td align="middle" ><font class="px12" color="black"><%=objrst("ze")-objrst("je")%></td>
                <td align="middle" ><font class="px12" color="blue"><%=formatpercent(objrst("bl"),2,-1)%></td>                                                                                                                                
                
                </tr>     
           <%objrst.Close %>     
                
             
             </tbody></table></td></tr></tbody></table>
       
             
            
<%end if%>    
  
 
   <br>
    
   
    </td>
  </tr>
</table>
</body>
</html>
