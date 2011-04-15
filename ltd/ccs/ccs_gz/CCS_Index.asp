<%@ Language=VBScript %>
<%
'Response.Write session("UID")
'Response.End
'session("UID")="{10EE558B-D8D1-11D4-8659-00805F594010}"

IF trim(session("UID"))<>"" then
  dim objD
   set ObjD=server.CreateObject ("Com_UserManage.ClsUserManage")
       VerifyOk_LDCX=objD.VerifyUserFunction (session("UID"),"CCS_LDCX")       
       VerifyOk_GSFZCX=objD.VerifyUserFunction (session("UID"),"CCS_GSFZCX") 
       VerifyOk_BMLDCX=objD.VerifyUserFunction (session("UID"),"CCS_BMLDCX")
       VerifyOk_GSGLY=objD.VerifyUserFunction (session("UID"),"CCS_GSGLY")
       VerifyOk_BMGLY=objD.VerifyUserFunction (session("UID"),"CCS_BMGLY")
       VerifyOk_GSCN=objD.VerifyUserFunction (session("UID"),"CCS_GSCN")
       VerifyOk_YTTS=objD.VerifyUserFunction (session("UID"),"CCS_CWYTTS")
       VerifyOk_GSCXY=objD.VerifyUserFunction (session("UID"),"CCS_Gscxy")
   
   if VerifyOk_LDCX=true then
     Response.Redirect "ccs_gsldcx_main.asp"  
   end if  
   if VerifyOk_GSFZCX=true then
     Response.Redirect "ccs_gsfzcx_main.asp"  
   end if      
   if VerifyOk_BMLDCX=true then     
     Response.Redirect "ccs_bmldcx_main.asp"  
   end if 
   if VerifyOk_GSGLY=true then     
     Response.Redirect "kmgl.asp"  
   end if 
   
   if VerifyOk_BMGLY=true then
     if month(date)<=1 then
   Response.Redirect "asppage1.asp"
   else
     'Response.Write "ccs_input_index.asp"
     'Response.End
     Response.Redirect "ccs_input_index.asp"
     end if  
   end if 
    
   if VerifyOk_GSCN=true then
     Response.Redirect "cwmain.asp"  
   end if    
   if VerifyOk_YTTS=true then
     Response.Redirect "ytgl.asp" 
   end if    
   if VerifyOk_GSCXY=true then
     Response.Redirect "ccs_gscxy_main.asp" 
   end if       
   
   
   session("errorNo")="000002"
   Response.Redirect "../sorry/sorry.asp"
   
else
   session("errorNo")="000001"
   Response.Redirect "../sorry/sorry.asp"
end if 
%>


