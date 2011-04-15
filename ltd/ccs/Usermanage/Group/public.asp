<% 
  if request.session("loginid") =""  then
   response.redirect "login.htm"
  end if
  
%> 