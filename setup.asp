<!-- #include file="function.asp" -->
<!-- #include file="cookie.asp" --><%
myfilename="setup.asp"
thisttime=Now()
if action="editsetup" and Request.ServerVariables("request_method") = "POST" then
filtrate=split(Request("badip"),"|")
for i = 0 to ubound(filtrate)
if instr("|"&remoteaddr&"","|"&filtrate(i)&"") > 0 then error("请检查您输入的IP地址是否正确")
next
set rs=server.createobject("ADODB.Recordset")
rs.Open "forumconfig",Conn,1,3
rs("forumname")=Request("eforumname")
rs("adduseryn")=Request("eadduseryn")
rs("RegTime")=Request("eRegTime")
rs("PostTime")=Request("ePostTime")
rs("contentlen")=Request("econtentlen")
rs("perpage")=Request("eperpage")
rs("wmhtmlkzn")=Request("ewmhtmlkzn")
rs("badwords")=Request("ebadwords")
rs("badip")=Request("ebadip")
rs("ifmodifytime")=Request("eifmodifytime")
rs.update
rs.close
set rs=nothing
message="<li>系统参数修改成功<li><a href=main.asp>返回管理首页</a>"
succeed(""&message&"<meta http-equiv=refresh content=2;url="&myfilename&">")
Response.End
end if''//##
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>系统设置</title>
<link href="images/css.css" rel="stylesheet" type="text/css">
</head>

<body>

<form name="form" action="<%=myfilename%>" method="post">
	<input type="hidden" name="action" value="editsetup">
	<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#C0C0C0" align="center">
		<tr bgcolor="#EEEEEE">
			<td align="right" width="25%" bgcolor="#FFFFFF">论坛名称：</td>
			<td bgcolor="#FFFFFF">
			<input name="eforumname" type="text" value="<%=forumname%>"></td>
		</tr>
		<tr bgcolor="#EEEEEE">
			<td align="right" bgcolor="#FFFFFF">是否允许注册用户：</td>
			<td bgcolor="#FFFFFF"><select name="eadduseryn"><%
	dim adduserynselect
	adduserynselect="<option value=0 >允许</option><option value=1 >不允许</option>"
	adduserynselect=replace(adduserynselect,"value="&adduseryn,"value="&adduseryn&" selected")
	%> <%=adduserynselect%></select></td>
		</tr>
		<tr bgcolor="#EEEEEE">
			<td align="right" bgcolor="#FFFFFF">用户注册间隔时间(秒)：</td>
			<td bgcolor="#FFFFFF">
			<input name="eRegTime" type="text" value="<%=RegTime%>" size="8"> 如果设不间隔就填0</td>
		</tr>
		<tr bgcolor="#EEEEEE">
			<td align="right" bgcolor="#FFFFFF">发帖间隔时间(秒)：</td>
			<td bgcolor="#FFFFFF">
			<input name="ePostTime" type="text" value="<%=PostTime%>" size="8"> 
			如果不防水就设为0</td>
		</tr>
		<tr bgcolor="#EEEEEE">
			<td align="right" bgcolor="#FFFFFF">限定内容的长度：</td>
			<td bgcolor="#FFFFFF">
			<input name="econtentlen" type="text" value="<%=contentlen%>" size="8"> 
			字节</td>
		</tr>
		<tr bgcolor="#EEEEEE">
			<td align="right" bgcolor="#FFFFFF">每页显示多少条主题：</td>
			<td bgcolor="#FFFFFF">
			<input name="eperpage" type="text" value="<%=perpage%>" size="8"></td>
		</tr>
		<tr bgcolor="#EEEEEE">
			<td align="right" bgcolor="#FFFFFF">文件扩展名：</td>
			<td bgcolor="#FFFFFF"><select name="ewmhtmlkzn" class="input1"><%
	dim wmhtmlkznselect
	wmhtmlkznselect="<option value=.html selected>.html</option><option value=.shtml>.shtml</option><option value=.xhtml>.xhtml</option>"
	wmhtmlkznselect=replace(wmhtmlkznselect,"value="&wmhtmlkzn,"value="&wmhtmlkzn&" selected")
	%> <%=wmhtmlkznselect%></select> 安装时设定后请不要改再动，否则之前的帖子因扩展名不同而无法查看</td>
		</tr>
		<tr bgcolor="#EEEEEE">
			<td align="right" bgcolor="#FFFFFF">修改后时否显示修改者及时间：</td>
			<td bgcolor="#FFFFFF"><select name="eifmodifytime" class="input1"><%
	dim ifmodifytimeselect
	ifmodifytimeselect="<option value=0>不显示</option><option value=1>显示</option>"
	ifmodifytimeselect=replace(ifmodifytimeselect,"value="&ifmodifytime,"value="&ifmodifytime&" selected")
	%> <%=ifmodifytimeselect%></select> </td>
		</tr>
		<tr bgcolor="#EEEEEE">
			<td align="right" bgcolor="#FFFFFF">不良词汇：</td>
			<td bgcolor="#FFFFFF">
			<input name="ebadwords" type="text" value="<%=badwords%>" size="40"> 
			用单竖线&quot;|&quot;隔开</td>
		</tr>
		<tr bgcolor="#EEEEEE">
			<td align="right" bgcolor="#FFFFFF">禁止IP：</td>
			<td bgcolor="#FFFFFF">
			<input name="ebadip" type="text" value="<%=badip%>" size="40"> 用单竖线&quot;|&quot;隔开</td>
		</tr>
		<tr bgcolor="#EEEEEE">
			<td bgcolor="#FFFFFF">　</td>
			<td bgcolor="#FFFFFF"><input type="submit" name="Submit" value="提交">
			<input type="reset" name="Submit" value="重置"></td>
		</tr>
	</table>
</form>

</body>

</html>