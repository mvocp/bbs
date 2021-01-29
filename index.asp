<!-- #include file="function.asp" --><%
myfilename="index.asp"
if action="setpage" then
if int(Request("pagesetup"))<>empty or int(Request("pagesetup"))<>0 then
Response.Cookies("pagesetup")=int(Request("pagesetup"))
else
Response.Cookies("pagesetup")=""
end if
end if
if action="out" and Request.ServerVariables("request_method") = "GET" then
Response.Cookies("username")=""
Response.Cookies("userpass")=""
message="<li>退出成功<li><a href=./>返回首页</a>"
succeed(""&message&"<meta http-equiv=refresh content=2;url=./>")
Response.End
end if
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=forumname%></title>
<link href="images/css.css" rel="stylesheet" type="text/css">
</head>

<body>

<table width="960" border="0" align="center" cellpadding="0" cellspacing="0">
	<tr>
		<td id="topads" align="center" colspan="3">广告数据加载中</td>
	</tr>
	<tr>
	<font color=#000000 size="5"><a href="/">返回首页</a> | <a href="newadd.asp">发表主题</a> 
		| <a href="?good=goodtopic">精华主题</a> | <%
if Request.Cookies("username")=empty or Request.Cookies("userpass")=empty then
%><a href="reg.asp">用户注册</a> | <a href="ulogin.asp">用户登陆</a> | <%
end if
%><a href="useredit.asp">修改资料</a> | <a href="?action=out">退出管理</font></a> <font color="#0000ff">(统计信息：共有主题<b><font color="#FF0000"><%=conn.execute("Select count(id)from [list] where upid=0")(0)%></font></b>条 回复<b><font color="#FF0000"><%=conn.execute("Select count(id)from [list] where upid>0")(0)%></font></b>条 注册会员<font color="#FF0000"><b><%=conn.execute("Select count(id)from [user]")(0)%></font></b>位)</font><hr site="0"></td>
	</tr>
	<%
dim topsql, pagesetup, count, TotalPage, PageCount, newtopic, rs2, sql2, replydata
dim wheresearch, goodsearch
set rs=server.createobject("ADODB.Recordset")
searchkey=HTMLEncode(Request("searchkey"))
wheresearch=""
if searchkey<>empty then wheresearch=" and txttitle like '%"&searchkey&"%' or username='"&searchkey&"'"
goodsearch=""
if Request("good")="goodtopic" then goodsearch=" and goodtopic=1"
if Request.Cookies("pagesetup")=empty then
pagesetup=perpage
else
pagesetup=int(Request.Cookies("pagesetup"))
if pagesetup > 150 then pagesetup=perpage
end if
topsql="where upid=0"&goodsearch&wheresearch&""
count=conn.execute("Select count(id) from [list] "&topsql&"")(0)
TotalPage=cint(count/pagesetup)
if TotalPage < count/pagesetup then TotalPage=TotalPage+1
PageCount = cint(Request.QueryString("P"))
if PageCount < 1 or PageCount = empty then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
sql="select * from [list] "&topsql&" order by toptopic Desc,lasttime Desc"
if PageCount>100 then
rs.Open sql,Conn,1
else
Set Rs=Conn.Execute(sql)
end if
if TotalPage>1 then RS.Move (PageCount-1) * pagesetup
i=0
Do While Not RS.EOF and i<pagesetup
i=i+1
if Not Response.IsClientConnected then responseend
list=list & RS("topictype") & ","
RS.MoveNext
loop
RS.Close
outmsg=""
if list<>empty then
	sql="select id,upid,icon,txttitle,username,ifchild,posttime,count,topictype,goodtopic,toptopic from [list] where topictype in ("&list&") order by toptopic desc,lasttime desc,orderu"
	rs.open sql,conn,1,1
	Do while (rs.eof=false)
	newtopic=""
	if rs("posttime")+1/24>now() then newtopic="<img src=images/new.gif>"
	if rs("goodtopic")=1 then
		topicimg="<img src=images/jinghua.gif>"
	elseif rs("toptopic")=1 then
		topicimg="<img src=images/top.gif>"
	else
		topicimg="<img src=images/icon.gif>"
	end if
	if rs("upid")=0 then
	response.write "</ul></td></tr>"&vbCrlf&"<tr><td><ul>"&topicimg&" <a style=font-size:14px; href=html/"&rs("id")&"/"&rs("id")&wmhtmlkzn&" target=_blank>"&rs("txttitle")&"</a> (<a href=reply.asp?upid="&rs("id")&">回<font color=#ff0000><b>"&rs("ifchild")&"</b></font>帖</a>) "&newtopic&" <font color=#888888>【"&rs("username")&"】 发表时间："&rs("posttime")&" 点击：<font color=#ff0000>"&rs("count")&"</font></font>"&vbCrlf&""
	else
	response.write "<ul><img src=images/reply.gif> <a href=html/"&rs("upid")&"/"&rs("id")&wmhtmlkzn&" target=_blank>"&rs("txttitle")&"</a> "&newtopic&"<font color=#888888>【"&rs("username")&"】 发表时间："&rs("posttime")&"</font></ul>"&vbCrlf&""
	end if
	rs.MoveNext
	loop
	rs.Close
end if
%><%=outmsg%>
</table>
<table width="960" border="0" cellpadding="1" cellspacing="0" align="center">
	<tr>
		<td bgcolor="#000000" width="15%" align="center" style="color: #FFFFFF">
		<%
dim pageselect, selectpagesetup
selectpagesetup=Request.Cookies("pagesetup")
if selectpagesetup=empty then selectpagesetup=0
pageselect="<option value=?action=setpage&pagesetup=30>30</option><option value=?action=setpage&pagesetup=10>10</option><option value=?action=setpage&pagesetup=30>30</option><option value=?action=setpage&pagesetup=40>40</option><option value=?action=setpage&pagesetup=50>50</option><option value=?action=setpage&pagesetup=80>80</option>"
pageselect=replace(pageselect,"pagesetup="&selectpagesetup,"pagesetup="&selectpagesetup&" selected")
%>每页显示 <select name="admnewstype" onchange="if(this.value!='no'){location.href=this.value}">
		<%=pageselect%></select> 条主题</td>
		<form name="form" action="<%=myfilename%>" method="post">
			<td align="center" bgcolor="#000000" style="color: #FFFFFF">内容搜索：<input name="searchkey" size="20" value="<%=Request("searchkey")%>">
			<input type="submit" value="搜索" name="submit"></td>
		</form>
		<td bgcolor="#000000" width="25%" style="color: #FFFFFF" align="center">
		本站共有 <%=TotalPage%> 页 [
		<script> 
TotalPage=<%=TotalPage%>
PageCount=<%=PageCount%>
for (var i=1; i <= TotalPage; i++) {
if (i <= PageCount+3 && i >= PageCount-3 || i==1 || i==TotalPage){
if (i > PageCount+4 || i < PageCount-2 && i!=1 && i!=2 ){document.write(" ... ");}
if (PageCount==i){document.write(" "+ i +" ");}
else{
document.write("<a href=?p="+i+"&good=<%=Request("good")%>&searchkey=<%=Request("searchkey")%>>"+ i +"</a> ");
}
}
}
</script>
		]</td>
	</tr>
	<tr>
		<td id="bottomads" align="center" colspan="3" height="40px">数据加载中</td>
	</tr>
	<tr>
		<td align="center" colspan="3">Copyright &copy; 2010 www.360-com.com Powered By <a href="http://www.360-com.com" target="_blank"><%=forumname%> v2.0</a></td>
	</tr>
</table>
<script language="JavaScript" src="images/topads.js"></script>
<script language="JavaScript" src="images/bottomads.js"></script>

</body>

</html>
<span style="display:none">
<script language="JavaScript" src="images/tj.js"></script>
</span>