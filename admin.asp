<!-- #include file="function.asp" -->
<!-- #include file="md5.asp" --><%
myfilename="admin.asp"
if action="userlogin" and Request.ServerVariables("request_method") = "POST" then
numcode=Request.Form("numcode")
username=HTMLEncode(Trim(Request.Form("login_username")))
userpass=md5(Trim(Request.Form("login_userpass")))
period=int(Request.Form("period"))
if username=empty or userpass=empty then error("����д����")
set rs=server.createobject("ADODB.Recordset")
If conn.Execute("Select id From [user] where distinction=10 and username='"&username&"' and userpass='"&userpass&"' " ).eof Then error("����д���û������������߲���վ��")

Response.Cookies("adminusername")=username
Response.Cookies("adminuserpass")=userpass
if period<>0 then
Response.Cookies("adminusername").Expires=date+period
Response.Cookies("adminuserpass").Expires=date+period
end if
message="<li>��½�ɹ�<li><a href=./admin1.asp>�����������</a>"
succeed(""&message&"<meta http-equiv=refresh content=2;url=./admin1.asp>")
Response.End
end if
if action="out" and Request.ServerVariables("request_method") = "GET" then
Response.Cookies("adminusername")=""
Response.Cookies("adminuserpass")=""
message="<li>�˳��������ĳɹ�<li><a href=./>������ҳ</a>"
succeed(""&message&"<meta http-equiv=refresh content=2;url=./>")
Response.End
end if
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��̨��½</title>
<link href="images/css.css" rel="stylesheet" type="text/css">
</head>

<body>

<form name="form" onsubmit="return VerifyInput();" action="<%=myfilename%>" method="post">
	<input type="hidden" name="action" value="userlogin">
	<table border="0" width="60%" cellspacing="1" cellpadding="3" bgcolor="#C0C0C0" align="center">
		<tr>
			<td bgcolor="#FFFFFF" align="right">�û�����</td>
			<td bgcolor="#FFFFFF">
			<input name="login_username" type="text" maxlength="15"></td>
		</tr>
		<tr>
			<td bgcolor="#FFFFFF" align="right">���룺</td>
			<td bgcolor="#FFFFFF"><input type="password" name="login_userpass"></td>
		</tr>
		<tr>
			<td bgcolor="#FFFFFF" align="right">��Ч�ڣ�</td>
			<td bgcolor="#FFFFFF"><select name="period">
			<option value="0" selected>������</option>
			<option value="1">һ��</option>
			<option value="30">һ��</option>
			<option value="360">����</option>
			</select> ���õ��Խ��鲻����</td>
		</tr>
		<tr>
			<td bgcolor="#FFFFFF" align="center" colspan="2">
			<input type="submit" name="button" value="�� ½"></td>
		</tr>
	</table>
</form>

</body>

</html>