<!-- #include file="function.asp" -->
<!-- #include file="cookie.asp" --><%
myfilename="topads.asp"
thisttime=Now()
dim etopads, outmsg
if action="topads" and Request.ServerVariables("request_method") = "POST" then
etopads=Request.form("etopads")
etopads=replace(etopads,";","")
etopads=replace(etopads,"'","")
sql="update topads set topads='"&etopads&"' where id=1"
conn.execute(sql)
outmsg=""
etopads=replace(etopads,vbCrlf,"")
outmsg="topads.innerHTML ='';"&vbCrlf&"topads.innerHTML +='"&etopads&"';"&vbCrlf&"topads.innerHTML +='';"
set fileobject = Server.CreateObject("Scripting.FileSystemObject")
set addhtmlfile = fileobject.CreateTextFile(Server.MapPath(".")&"\images\topads.js")
addhtmlfile.writeline outmsg

message="<li>顶部广告修改成功<li><a href=main.asp>返回管理首页</a>"
succeed(""&message&"<meta http-equiv=refresh content=2;url="&myfilename&">")
Response.End
end if
Set Rs=Conn.Execute("topads")
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>顶部广告设置</title>
<link href="images/css.css" rel="stylesheet" type="text/css">
</head>

<body>

<form name="FORM" action="<%=myfilename%>" method="post">
	<input type="hidden" name="action" value="topads">
	<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#C0C0C0" align="center">
		<tr>
			<td align="center" bgcolor="#FFFFFF">
			<textarea name="etopads" cols="80" rows="20"><%=rs("topads")%></textarea></td>
		</tr>
		<tr>
			<td align="center" bgcolor="#FFFFFF">
			<input type="submit" name="Submit" value="提交">&nbsp;
			<input type="reset" name="Submit" value="重置">  [页面顶部支持html]</td>
		</tr>
	</table>
</form>

</body>

</html>