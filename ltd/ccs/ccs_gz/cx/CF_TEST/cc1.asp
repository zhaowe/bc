<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">






<SCRIPT LANGUAGE=javascript>
<!--
function changeColor()
{
  if (document.body.bgColor == "#ff0000")   // Check if body bgColor is red by comparing to hexidecimal value
      document.body.bgColor = "blue";
  else
      document.body.bgColor = "red";
}

function me()
{
setInterval("changeColor()", 1000*60*30);
}

//-->
</SCRIPT>







</HEAD>



<BODY onload="me();">

<P>&nbsp;</P>

</BODY>
</HTML>
