<html>
<link rel="stylesheet" href="../../style.css">
<body>
<%
on error resume next
dim userid
userid=session("UID")
dim password1,password2

password1=Session("Password1") 
password2=Session("Password2")


if lcase(trim(session("loginid")))<>lcase(trim(password1)) then
if password1<>"" then
   if len(cstr(password1))<=10 then
     if len(trim(password1))>=8 then
       if password1=password2 then
       
          
       '密码改成8-10的数字字母组合
       
        temppwd=password1
        temppwd=ucase(trim(temppwd))
        havechar=0
        havenum=0
        length=len(password1)
        
        
        
        for i=0 to length
          
          temppwd=left(temppwd,length-i)
          current=right(temppwd,1)
          if current>="0" and  current<="9" then
             havenum=1
          'else
          end if
          if current>="A" and  current<="Z" then
             havechar=1
          '    else
          '       Response.Write "密码不允许有除英文字母和数字以外的字符！"                 
          end if   
          'end if
         next
        
        
        
        
        
        if havenum+havechar=2 then
           
       
       '密码改成8-10的数字字母组合
       
       
      
      
          Set objConn = Server.CreateObject("ADODB.Connection")
          objConn.Open Application("OledbStr") 
      
          Set objRst=server.CreateObject ("ADODB.Recordset")
          objRst.LockType=3
          objRst.CursorType=3
          set objRst.activeConnection=objConn

          objRst.Source="select loginid,password from userinfo where loginid='"& trim(session("LoginID")) &"' and password='"& trim(password1) &"' "
          objRst.Open 
 
          if objrst.EOF and objrst.BOF then       
        '判断新旧密码是否一致
        
        
       
         dim objdml
         set objdml=server.CreateObject ("Com_UserManage.ClsUserManage")
         ierror=objdml.EditPassword (UserId,password1)
         if Err.number <> 0 then
	        ierror=Err.number
	        Err.Clear
	        set objdml=nothing 
	        Response.Redirect "../../Sorry.asp?Errorno="&ierror
         end if
         set objdml=nothing
         Response.Write "修改密码成功!"
         
        
        '新旧密码如果一样
         else
            Response.Write "新密码和原来旧密码不能一样！"      
         end if 
         
        
         else
            Response.Write "密码必须是英文字母与数字的组合！"                 
         end if
       

       
       
       else 
         Response.Write "对不起，您输入的新密码不正确，请点击<返回>按钮重新输入！"
       end if
      else
       Response.Write "对不起，您输入的新密码不能少于8位！" 
      end if
     else
       Response.Write "对不起，您输入的新密码不能多于10位！"
   end if
 else
   Response.Write "对不起，请再次输入您的新密码！您输入的新密码不能为空！"
end if
else
   Response.Write "对不起，密码和用户名不能完全一样！"
end if
%>
<p>[ <a href="javascript:history.back()">返回</a> ]</p>
</body>
</html>