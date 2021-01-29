<!-- #include file="function.asp" -->
<!-- #include file="cookie.asp" --><%
Dim theInstalledObjects(18)
dim duanke
theInstalledObjects(0) = "MSWC.AdRotator"
theInstalledObjects(1) = "MSWC.BrowserType"
theInstalledObjects(2) = "MSWC.NextLink"
theInstalledObjects(3) = "MSWC.Tools"
theInstalledObjects(4) = "MSWC.Status"
theInstalledObjects(5) = "MSWC.Counters"
theInstalledObjects(6) = "IISSample.ContentRotator"
theInstalledObjects(7) = "IISSample.PageCounter"
theInstalledObjects(8) = "MSWC.PermissionChecker"
theInstalledObjects(9) = "Scripting.FileSystemObject"
theInstalledObjects(10) = "adodb.connection"
theInstalledObjects(11) = "SoftArtisans.FileUp"
theInstalledObjects(12) = "SoftArtisans.FileManager"
theInstalledObjects(13) = "JMail.Message"
theInstalledObjects(14) = "CDONTS.NewMail"
theInstalledObjects(15) = "Persits.MailSender"
theInstalledObjects(16) = "LyfUpload.UploadFile"
theInstalledObjects(17) = "Persits.Upload.1"
theInstalledObjects(18) = "w3.upload"

duanke=""
if Request.ServerVariables("SERVER_PORT")<>80 then duanke=":"&Request.ServerVariables("SERVER_PORT")
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>系统信息</title>
<link href="images/css.css" rel="stylesheet" type="text/css">
</head>

<body>

<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#C0C0C0" align="center">
	<tr>
		<td width="50%" bgcolor="#FFFFFF">服务器域名：http://<%=Request.ServerVariables("server_name")%><%=duanke%></td>
		<td bgcolor="#FFFFFF">服务器操作系统：<%=Request.ServerVariables("OS")%></td>
	</tr>
	<tr>
		<td bgcolor="#FFFFFF">服务器软件的名称：<%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
		<td bgcolor="#FFFFFF">脚本解释引擎：<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
	</tr>
	<tr>
		<td bgcolor="#FFFFFF">FSO 文本读写：<%If Not IsObjInstalled(theInstalledObjects(9)) Then%><font color="red"><b>×</b></font><%else%><b>√</b><%end if%></td>
		<td bgcolor="#FFFFFF">数据库路径：<a target="_blank" href="<%=datapath%><%=datafile%>"><%=datapath%><%=datafile%></a></td>
	</tr>
	<tr>
		<form method="post" name="form" action="admin_list.asp">
			<td bgcolor="#FFFFFF">快捷查找帖子：<input type="text" name="searchkey">
			<input type="submit" name="Submit" value="查找"></td>
		</form>
		<form method="post" name="form" action="admin_member.asp">
			<input type="hidden" name="action" value="search">
			<td bgcolor="#FFFFFF">快捷查找用户：<input type="text" name="searchkey">
			<input type="submit" name="Submit" value="查找"></td>
		</form>
	</tr>
</table>
</body>

</html>
<%
Function IsObjInstalled(strClassString)
On Error Resume Next
IsObjInstalled = False
Err = 0
Dim xTestObj
Set xTestObj = Server.CreateObject(strClassString)
If 0 = Err Then IsObjInstalled = True
Set xTestObj = Nothing
Err = 0
End Function
%>