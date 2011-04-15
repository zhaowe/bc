<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script language="javascript1.2"> 
var tmonth=new Array("January","February","March","April","May","June","July","August","September","October","November","December");
var tday=new Array("Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"); function fixnum(num)
{
    if(num<10)
    num='0'+num;
    return num;
} 

function fixmd(name,n)
{
    var i=parseInt(n);
    var md=name[i];
    return md;
} 

function timerun()
{
var today=new Date();
var t_year=today.getYear();
var t_month=today.getMonth();
var t_date=today.getDate();
var t_day=today.getDay();
var t_hour=today.getHours();
var t_minute=today.getMinutes();
var t_second=today.getSeconds();
var t_apm;

    if(t_hour<12)
    t_apm="am";
    else
    {
    t_apm="pm";
    t_hour-=12;
    }
var minute_t=fixnum(t_minute);    
var second_t=fixnum(t_second);
var month_t=fixmd(tmonth,t_month);
var day_t=fixmd(tday,t_day); var the_date=t_year+" "+month_t+" "+t_date;
var the_time=t_hour+":"+minute_t+":"+second_t+"."+t_apm+" "+day_t;

i1.value=the_date;
i2.value=the_time;
setTimeout("timerun()",1000); 

}

 </script> 
</HEAD>
</SCRIPT>


<body onload="timerun();">
<input id=i1 style="font-size: 60pt;color:dodgerblue;border:0;width:100%" type="text"><br>
<input id=i2 style="font-size: 60pt;color:dodgerblue;border:0;width:100%" type="text"> 

</body> 
