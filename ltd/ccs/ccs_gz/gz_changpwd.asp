<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>修改密码</title>
</head>

<body>

  
 <p align="center">　 </p>


 

 <p align="center"><font face="宋体" size="6" color="#6666ff"><b>修改密码</b></font> </p>


<div align="center">
  <center>
  <form method="post" action="gz_changpwd.asp"  id=form1 name=form1>
  <table border="1" width="46%" bordercolor="#ff99ff">
  <tr>
    <td width="100%" align="middle"><font color="#6600ff">用户名：</font><input  name="loginid">
    </td>
  </tr>
  <tr>
    <td width="100%" align="middle"><font color="#6600ff">旧密码：</font><input type="password" name="password0" 
     ></td>
  </tr>
  <tr>
    <td width="100%" align="middle"><font color="#6600ff">新密码：</font><input type="password" name="password1" 
     ></td>
  </tr>
  <tr>
    <td width="100%" align="middle"><font color="#6600ff">再输入：</font><input type="password" name="password2" 
     ></td>
  </tr>
  <tr>
    <td width="100%" align="middle"><input type="submit" value="提交" name="B1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
      <input type="reset" value="取消" name="B2">
      <input type="hidden" value=1 name="tj">  
    </td> 
      
  </tr> 
  <tr>
    <td width="100%" align="middle"><font color=red> 密码必须是英文字母与数字的组合；<br> 8&lt;=密码长度&lt;=10；新密码不能与旧密码相同！ </font></td>
  </tr>  
</table> 
</form> 
<p>
[ <A href="javascript:history.back()">返回</A> ]
</p>
  </center> 
</div> 
 
<%
   Set objConn = Server.CreateObject("ADODB.Connection")
   objConn.Open Application("OledbStr") 
      
   Set objRst=server.CreateObject ("ADODB.Recordset")
   objRst.LockType=3
   objRst.CursorType=3
   set objRst.activeConnection=objConn
   tj=Request.Form("tj")
   
   if tj=1 then
   
dim loginid
loginid=Request.Form("loginid")
dim password0,password1,password2

password0=Request.Form("Password0") 
password1=Request.Form("Password1") 
password2=Request.Form("Password2")


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
       
        
      
                    
          objRst.Source="select loginid,password from userinfo where loginid='"& trim(LoginID) &"' and password='"& trim(password0) &"' "
          'Response.Write objrst.Source
          'Response.End
          objRst.Open 
          
          if not (objrst.EOF or objrst.BOF) then       
          
          objrst.Close   
          objRst.Source="select loginid,password from userinfo where loginid='"& trim(LoginID) &"' and password='"& trim(password1) &"' "
          objRst.Open 
 
          if objrst.EOF and objrst.BOF then       
        '判断新旧密码是否一致
        
        
         objrst.Close
         objrst.Source="update userinfo set password='" & trim(password1)  & "' where loginid='" & trim(LoginID) & "'"
         objrst.Open
         
         Response.Write("<script>window.alert('修改密码成功!')</script>") '"修改密码成功!"
         
        
        '新旧密码如果一样
         else
            Response.Write "新密码和原来旧密码不能一样！"      
         end if 
         
        else
            Response.Write "你输入了错误的旧密码或用户名！"                 
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
   
   end if
%>


</body> 
 
</html> 
