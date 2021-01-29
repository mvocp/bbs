<!-- #include file="function.asp" -->
<!-- #include file="cookie.asp" --><%
myfilename="list.asp"
thisttime=Now()
dim listid, listtype
set rs=server.createobject("ADODB.Recordset")
select case action
	case "topbbs"
		for each listcheckbox in Request.form("listcheckboxs")
			listid=""
			listtype=""
			checkbox=split(listcheckbox,"_")
			listid=int(HTMLEncode(Trim(checkbox(0))))
			listtype=int(HTMLEncode(Trim(checkbox(1))))
			sql="update list set toptopic=1 where topictype="&listtype
			conn.execute(sql)
		next
		Set rs=Nothing
		message="<li>文章置顶成功<li><a href=main.asp>返回管理首页</a>"
		succeed(""&message&"<meta http-equiv=refresh content=2;url="&myfilename&">")
		Response.End
	case "canceltop"
		for each listcheckbox in Request.form("listcheckboxs")
			listid=""
			listtype=""
			checkbox=split(listcheckbox,"_")
			listid=int(HTMLEncode(Trim(checkbox(0))))
			listtype=int(HTMLEncode(Trim(checkbox(1))))
			sql="update list set toptopic=0 where topictype="&listtype
			conn.execute(sql)
		next
		Set rs=Nothing
		message="<li>文章消顶成功<li><a href=main.asp>返回管理首页</a>"
		succeed(""&message&"<meta http-equiv=refresh content=2;url="&myfilename&">")
		Response.End
	case "lastbbs"
		for each listcheckbox in Request.form("listcheckboxs")
			listid=""
			listtype=""
			checkbox=split(listcheckbox,"_")
			listid=int(HTMLEncode(Trim(checkbox(0))))
			listtype=int(HTMLEncode(Trim(checkbox(1))))
			sql="update list set lasttime='"&thisttime&"' where topictype="&listtype
			conn.execute(sql)
		next
		Set rs=Nothing
		message="<li>文章拉前成功<li><a href=main.asp>返回管理首页</a>"
		succeed(""&message&"<meta http-equiv=refresh content=2;url="&myfilename&">")
		Response.End
	case "goodbbs"
		for each listcheckbox in Request.form("listcheckboxs")
			listid=""
			listtype=""
			checkbox=split(listcheckbox,"_")
			listid=int(HTMLEncode(Trim(checkbox(0))))
			listtype=int(HTMLEncode(Trim(checkbox(1))))
			sql="update list set goodtopic=1 where topictype="&listtype
			conn.execute(sql)
		next
		Set rs=Nothing
		message="<li>文章精华成功<li><a href=main.asp>返回管理首页</a>"
		succeed(""&message&"<meta http-equiv=refresh content=2;url="&myfilename&">")
		Response.End
	case "cancelgood"
		for each listcheckbox in Request.form("listcheckboxs")
			listid=""
			listtype=""
			checkbox=split(listcheckbox,"_")
			listid=int(HTMLEncode(Trim(checkbox(0))))
			listtype=int(HTMLEncode(Trim(checkbox(1))))
			sql="update list set goodtopic=0 where topictype="&listtype
			conn.execute(sql)
		next
		Set rs=Nothing
		message="<li>文章消精成功<li><a href=main.asp>返回管理首页</a>"
		succeed(""&message&"<meta http-equiv=refresh content=2;url="&myfilename&">")
		Response.End
	case "delbbs"
		set fileobject = Server.CreateObject("Scripting.FileSystemObject")
		for each listcheckbox in Request.form("listcheckboxs")
			listid=""
			listtype=""
			checkbox=split(listcheckbox,"_")
			listid=int(HTMLEncode(Trim(checkbox(0))))
			listtype=int(HTMLEncode(Trim(checkbox(1))))
			sql="delete from [list] where topictype="&listtype
			conn.execute(sql)
			if fileobject.FolderExists(Server.MapPath(".\html")&"\"&listid&"") then
				fileobject.DeleteFolder(Server.MapPath(".\html")&"\"&listid&"")
			end if
		next
		Set rs=Nothing
		message="<li>删除主题及回复成功<li><a href=main.asp>返回管理首页</a>"
		succeed(""&message&"<meta http-equiv=refresh content=2;url="&myfilename&">")
		Response.End
	case "delreply"
		set fileobject = Server.CreateObject("Scripting.FileSystemObject")
			sql="delete from [list] where id="&Request("replyid")
			conn.execute(sql)
			fileobject.DeleteFile(Server.MapPath(".\html")&"\"&Request("dupid")&"\"&Request("replyid")&wmhtmlkzn&"")
		Set rs=Nothing
		message="<li>删除回复成功<li><a href=main.asp>返回管理首页</a>"
		succeed(""&message&"<meta http-equiv=refresh content=2;url="&myfilename&">")
		Response.End
	case else
		action=""
end select

%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>内容管理</title>
<link href="images/css.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
	function delete_confirm(){
	if (confirm("您是否要删除选定的文章？")){
		GoSubmit('delbbs');
	}
	return false;
	}
	function CheckAll(form)
	  {
	  for (var i=0;i<form.elements.length;i++)
	    {
	    var e = form.elements[i];
	          e.checked = true;
	    }
	  }
	function FanAll(form)
	 {
	  for (var i=0;i<form.elements.length;i++)
	    {
	    var e = form.elements[i];
	      if (e.checked == true){
	          e.checked = false;
	          }
	       else {
	          e.checked = true;
	          }
	    }
	}
	function GoSubmit(action)
	{
		document.bbs.action.value = action;
		document.bbs.submit();
	}
	</script>
</head>

<body>

<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#C0C0C0" align="center">
	<tr>
		<td align="center" bgcolor="#FFFFFF" colspan="2">
		<input type="button" name value="刷新" onclick="location.reload()">
		<input type="button" name value="全选" onclick="CheckAll(bbs)">
		<input type="button" name value="反选" onclick="FanAll(bbs)">
		<input type="button" name value="重置" onclick="document.bbs.reset()">
		<input type="button" name value="置顶" onclick="GoSubmit('topbbs')">
		<input type="button" name value="消顶" onclick="GoSubmit('canceltop')">
		<input type="button" name value="拉前" onclick="GoSubmit('lastbbs')">
		<input type="button" name value="精华" onclick="GoSubmit('goodbbs')">
		<input type="button" name value="消精" onclick="GoSubmit('cancelgood')">
		<input type="button" name value="删除" onclick="return delete_confirm();"></td>
	</tr>
	<form name="bbs" method="POST" action="<%=myfilename%>">
		<input type="hidden" name="action" value><%
dim topsql, pagesetup, count, TotalPage, PageCount, newtopic, rs2, sql2, replydata
dim wheresearch, goodsearch
searchkey=HTMLEncode(Request("searchkey"))
wheresearch=""
if searchkey<>empty then wheresearch=" and txttitle like '%"&searchkey&"%'"
if Request("good")="goodtopic" then goodsearch=" and goodtopic=1"
if Request.Cookies("pagesetup")=empty then
pagesetup=perpage
else
pagesetup=int(Request.Cookies("pagesetup"))
if pagesetup > 150 then pagesetup=perpage
end if
topsql="where upid=0"&goodsearch&wheresearch&""
count=conn.execute("Select count(id) from [list] "&topsql&"")(0)
TotalPage=cint(count/pagesetup)  '总页数
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
	sql="select id,upid,icon,txttitle,username,posttime,count,topictype,goodtopic,toptopic from [list] where topictype in ("&list&") order by toptopic desc,lasttime desc,orderu"
	rs.open sql,conn,1,1
	Do while (rs.eof=false)
	topicimg=""
	if rs("goodtopic")=1 then
		topicimg="<img src=images/jinghua.gif>"
	elseif rs("toptopic")=1 then
		topicimg="<img src=images/top.gif>"
	else
		topicimg="<img src=images/icon.gif>"
	end if
	if rs("upid")=0 then
	response.write "</ul></td></tr>"&vbCrlf&"<tr><td bgcolor=#FFFFFF colspan=2><ul><input type=checkbox name=listcheckboxs value="&rs("id")&"_"&rs("topictype")&"> "&topicimg&" <a href=html/"&rs("id")&"/"&rs("id")&wmhtmlkzn&" target=_blank>"&rs("txttitle")&"</a> <font color=#888>(<font color=0000ff>"&rs("username")&"</font> 发表时间："&rs("posttime")&" 点击：<font color=#ff0000>"&rs("count")&"</font>)</font>"&vbCrlf&""
	else
	response.write "<ul><img src=images/reply.gif> <a href=html/"&rs("upid")&"/"&rs("id")&wmhtmlkzn&" target=_blank>"&rs("txttitle")&"</a> <font color=#999>(<font color=0000ff>"&rs("username")&"</font> 发表时间："&rs("posttime")&" 点击：<font color=red>"&rs("count")&"</font>) <a href="&myfilename&"?action=delreply&replyid="&rs("id")&"&dupid="&rs("upid")&">删除该回复</a></ul></font>"&vbCrlf&""
	end if
	rs.MoveNext
	loop
	rs.Close
end if
%><%=outmsg%>
	</form>
	<tr>
		<td bgcolor="#FFFFFF">本论坛共有 <%=TotalPage%> 页 [
		<script>
TotalPage=<%=TotalPage%>
PageCount=<%=PageCount%>
for (var i=1; i <= TotalPage; i++) {
if (i <= PageCount+3 && i >= PageCount-3 || i==1 || i==TotalPage){
if (i > PageCount+4 || i < PageCount-2 && i!=1 && i!=2 ){document.write(" ... ");}
if (PageCount==i){document.write(" "+ i +" ");}
else{
document.write("<a href=?p="+i+"&good=<%=Request("good")%>&searchkey=<%=Request("searchkey")%>><font color=#ffffff>"+ i +"</font></a> ");
}
}
}
</script>
		]</td>
		<form name="form" action="<%=myfilename%>" method="post">
			<td width="40%" align="right" bgcolor="#FFFFFF">快速搜索：<input name="searchkey" value="<%=Request("searchkey")%>" size="20">
			<input type="submit" value="搜索" name="submit"></td>
		</form>
	</tr>
</table>

</body>

</html>