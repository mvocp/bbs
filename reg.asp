<!-- #include file="function.asp" -->
<!-- #include file="MD5.asp" --><%
myfilename="reg.asp"

if adduseryn = 1 then error("本论坛暂时不开放新用户注册！")
if action="adduser" and Request.ServerVariables("request_method") = "POST" then
if Request.Cookies("regusertime")<>empty then
if DateDiff("s",Request.Cookies("regusertime"),Now()) < RegTime then error("论坛限制两次注册时间间隔"&RegTime&"秒！")
end if
username=HTMLEncode(Trim(Request.Form("username")))
errorchar=array(" ","　","	","#","`","|","%","&","","+",";")
for i=0 to ubound(errorchar)
if instr(username,errorchar(i))>0 then error("用户名中不能含有特殊符号")
next

userpass=Trim(Request.Form("userpass"))
userrpass=Trim(Request.Form("userrpass"))
userface=HTMLEncode(Request.Form("userhead"))
if username="" then  error("您的用户名没有填写")
if Len(username)<3 then error("您的用户名中不能少于3字节")
if Len(username)>15 then error("您的用户名中不能超过15字节")
set rs=server.createobject("ADODB.Recordset")
If not conn.Execute("Select id From [user] where username='"&username&"'" ).eof Then error("您填写的用户已有人注册了")
if userpass=empty then error("您没有填写密码")
if Len(userpass)<6 then error("您的密码不能少于6字节")
if userpass<>userrpass then error("您两次输入的密码不相同")
if instr(userface,";")>0 then error("您的头像选择有错")

rs.Open "[user]",conn,1,3
rs.addnew
rs("username")=username
rs("userpass")=md5(userpass)
rs("userface")=userface
rs("userip")=remoteaddr
rs.update
rs.close
Set rs=Nothing

Response.Cookies("username")=username
Response.Cookies("userpass")=md5(userpass)
Response.Cookies("regusertime")=Now()
message="<li>用户注册成功<li><a href=./>返回论坛首页</a>"
succeed(""&message&"<meta http-equiv=refresh content=2;url=./>")

Response.End
end if'''''''''''
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=forumname%> - 用户注册</title>
<link href="images/css.css" rel="stylesheet" type="text/css">
</head>

<body>

<form name="form" onsubmit="return VerifyInput();" action="<%=myfilename%>" method="post">
	<input type="hidden" name="action" value="adduser">
	<table border="0" width="60%" cellspacing="1" cellpadding="3" bgcolor="#C0C0C0" align="center">
		<tr>
			<td bgcolor="#FFFFFF" rowspan="5" align="center">
			<img src="images/face/1.gif" name="tus"></td>
			<td bgcolor="#FFFFFF" align="right">用户名：</td>
			<td bgcolor="#FFFFFF">
			<input name="username" type="text" size="20" maxlength="24"></td>
		</tr>
		<tr>
			<td bgcolor="#FFFFFF" align="right">密码：</td>
			<td bgcolor="#FFFFFF">
			<input name="userpass" type="password" size="20" maxlength="20"></td>
		</tr>
		<tr>
			<td bgcolor="#FFFFFF" align="right">确认密码：</td>
			<td bgcolor="#FFFFFF">
			<input name="userrpass" type="password" size="20" maxlength="20"></td>
		</tr>
		<tr>
			<td bgcolor="#FFFFFF" align="right">头像设置：</td>
			<td bgcolor="#FFFFFF">
			<script>function showimage(){document.images.tus.src="images/face/"+document.form.userhead.options[document.form.userhead.selectedIndex].value+".gif";}</script>
			请选择：<select name="userhead" size="1" onchange="showimage()">
			<option value="1">1</option>
			<option value="2">2</option>
			<option value="3">3</option>
			<option value="4">4</option>
			<option value="5">5</option>
			<option value="6">6</option>
			<option value="7">7</option>
			<option value="8">8</option>
			<option value="9">9</option>
			<option value="10">10</option>
			<option value="11">11</option>
			<option value="12">12</option>
			<option value="13">13</option>
			<option value="14">14</option>
			<option value="15">15</option>
			<option value="16">16</option>
			<option value="17">17</option>
			<option value="18">18</option>
			<option value="19">19</option>
			<option value="20">20</option>
			<option value="21">21</option>
			<option value="22">22</option>
			<option value="23">23</option>
			<option value="24">24</option>
			</select></td>
		</tr>
		<tr>
			<td bgcolor="#FFFFFF" align="center" colspan="2">
			<input type="submit" name="button" value="注 册"> <a href="/">【返回首页】</a></td>
		</tr>
	</table>
</form>

</body>

</html>
<script>
function VerifyInput()
{
username=document.form.username.value

if (username == "")
{
alert("请输入您的用户名");
document.form.username.focus();
return false;
}
return true;
}
</script>
