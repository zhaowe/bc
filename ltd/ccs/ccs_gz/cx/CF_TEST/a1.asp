<%@ Language=VBScript %>
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>





<SCRIPT  LANGUAGE=javascript>
<!--

function button1_onclick() {
document.bgColor ="green"
}

//-->
</SCRIPT>



<SCRIPT LANGUAGE=vbscript>
<!--

Sub submit1click

'call window_onunload
 window.alert (gg)
 
sn_value=trim(form1.sn.value)
if isnumeric(sn_value) then
   if sn_value<99999 then
      window.alert ("是数字")
      '进行网页提交
      form1.submit ()
      
   else
      window.alert (sn_value+"数字太大了哦")
      'form1.sn.focus()
   end if   
else
   window.alert ("不是数字哦")
   'form1.sn.focus()       
end if

End Sub

-->
</SCRIPT>








<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

Sub window_onunload
   form1.sn.value ="9999999999999999"
   document.bgColor ="red"   
End Sub









Sub sn_onfocus
'window.confirm( form1.sn.value )
gg=form1.sn.value 
End Sub

-->
</SCRIPT>



</HEAD>


<BODY LANGUAGE=vbscript onunload=" window_onunload()">
<P>&nbsp;</P>
<form id=form1 name=form1>

<P><INPUT id=sn name=sn value="888888"></P>
<P>&nbsp;</P>
<P>&nbsp;</P>
<P><INPUT type=button value=button name=button1 LANGUAGE=javascript onclick="return button1_onclick()"></P>






<P>&nbsp;</P>
<P><INPUT type=button value=Submit name=submit1 LANGUAGE=vbscript onclick="submit1click()"></P>

</form>

</BODY>
</HTML>
