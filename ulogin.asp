<!-- #include file="function.asp" -->
<!-- #include file="MD5.asp" --><%
myfilename="ulogin.asp"
if action="userlogin" and Request.ServerVariables("request_method") = "POST" then
username=HTMLEncode(Trim(Request.Form("login_username")))
userpass=md5(Trim(Request.Form("login_userpass")))
period=int(Request.Form("period"))
if username=empty or userpass=empty then error("请填写完整")
set rs=server.createobject("ADODB.Recordset")
If conn.Execute("Select id From [user] where username='"&username&"' and userpass='"&userpass&"' " ).eof Then error("你填写的用户名或者密码错误")

Response.Cookies("username")=username
Response.Cookies("userpass")=userpass
if period<>0 then
Response.Cookies("username").Expires=date+period
Response.Cookies("userpass").Expires=date+period
end if
message="<li>登陆成功<li><a href=./>返回论坛首页</a>"
succeed(""&message&"<meta http-equiv=refresh content=2;url=./>")
Response.End
end if
if action="outlogin" and Request.ServerVariables("request_method") = "GET" then
Response.Cookies("username")=""
Response.Cookies("userpass")=""
message="<li>退出论坛成功<li><a href=./>返回论坛首页</a>"
succeed(""&message&"<meta http-equiv=refresh content=2;url=./>")
Response.End
end if

%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=forumname%> - 用户登陆</title>
<link href="images/css.css" rel="stylesheet" type="text/css">
</head>

<body>

<form name="form" onsubmit="return VerifyInput();" action="<%=myfilename%>" method="post">
	<input type="hidden" name="action" value="userlogin">
	<table border="0" width="60%" cellspacing="1" cellpadding="3" bgcolor="#C0C0C0" align="center">
		<tr>
			<td bgcolor="#FFFFFF" align="right">用户名：</td>
			<td bgcolor="#FFFFFF">
			<input name="login_username" type="text" maxlength="24"></td>
		</tr>
		<tr>
			<td bgcolor="#FFFFFF" align="right">密码：</td>
			<td bgcolor="#FFFFFF"><input type="password" name="login_userpass"></td>
		</tr>
		<tr>
			<td bgcolor="#FFFFFF" align="right">有效期：</td>
			<td bgcolor="#FFFFFF"><select name="period">
			<option value="0" selected>不保存</option>
			<option value="1">一天</option>
			<option value="30">一月</option>
			<option value="360">永久</option>
			</select> 公用电脑建议不保存</td>
		</tr>
		<tr>
			<td bgcolor="#FFFFFF" colspan="2" align="center">
			<input type="submit" name="button" value="登 陆"> <a href="/">【返回首页】</a></td>
		</tr>
	</table>
</form>

</body>

</html>
