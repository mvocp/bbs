<!-- #include file="function.asp" -->
<!-- #include file="cookie.asp" --><%
myfilename="bottomads.asp"
thisttime=Now()
dim ebottomads, outmsg
if action="bottomads" and Request.ServerVariables("request_method") = "POST" then
ebottomads=Request.form("ebottomads")
ebottomads=replace(ebottomads,";","")
ebottomads=replace(ebottomads,"'","")
sql="update bottomads set bottomads='"&ebottomads&"' where id=1"
conn.execute(sql)
outmsg=""
ebottomads=replace(ebottomads,vbCrlf,"")
outmsg="bottomads.innerHTML ='';"&vbCrlf&"bottomads.innerHTML +='"&ebottomads&"';"&vbCrlf&"bottomads.innerHTML +='';"
set fileobject = Server.CreateObject("Scripting.FileSystemObject")
set addhtmlfile = fileobject.CreateTextFile(Server.MapPath(".")&"\images\bottomads.js")
addhtmlfile.writeline outmsg

message="<li>ҳ�Ź���޸ĳɹ�<li><a href=main.asp>���ع�����ҳ</a>"
succeed(""&message&"<meta http-equiv=refresh content=2;url="&myfilename&">")
Response.End
end if
Set Rs=Conn.Execute("bottomads")
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>ҳ�׹��</title>
<link href="images/css.css" rel="stylesheet" type="text/css">
</head>

<body>

<form name="FORM" action="<%=myfilename%>" method="post">
	<input type="hidden" name="action" value="bottomads">
	<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#C0C0C0" align="center">
		<tr>
			<td align="center" bgcolor="#FFFFFF">
			<textarea name="ebottomads" cols="80" rows="20"><%=rs("bottomads")%></textarea></td>
		</tr>
		<tr>
			<td align="center" bgcolor="#FFFFFF">
			<input type="submit" name="Submit" value="�ύ">&nbsp;
			<input type="reset" name="Submit" value="����"> [ҳ��ײ����֧��html]</td>
		</tr>
	</table>
</form>

</body>

</html>