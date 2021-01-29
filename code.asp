<%
function html_encode(byval s)
	dim hl1,hl2
	if isnull(s) then
		html_encode=""
	else
		s=replace(s,"<","&lt;")
		s=replace(s,">","&gt;")
		s=replace(s,chr(32)," ")
		s=replace(s,chr(9), " ")
		s=replace(s,chr(39),"&#39;")
		s=replace(s,chr(34),"&quot;")
		s=replace(s,chr(13),"")
  	s=replace(s,chr(10),"<br>")
  	html_encode = s
	end if
end function
'----------------------------------------------------
function cnum(s)
	if isnull(s) or s=""  then
		exit function
	else
		if not isnumeric(s) then
			response.write"<center>非法操作导致程序中止!</center>"
			response.end
		else
			cnum=int(s)
		end if
	end if
end function
'----------------------------------------------------
function text_encode(byval str)
	if isnull(str) then
		text_encode=""
	else
		str=replace(str,"&","&amp;")
		str=replace(str,"<","&lt;")
		text_encode=replace(str,">","&gt;")
	end if
end function
'-----------------------------------------------------
dim re
function ubb(byval s)
	if isnull(s) or s="" then
	  ubb = ""
	else
		ubb = ubb_code(s)
	end if
end function
'--------------------------------------------------------
function ubbcode(s,uCodeL,uCodeR,tCode)
	re.Pattern=uCodeL&uCodeR
	s=re.Replace(s,"")
	re.Pattern=uCodeL&"(.+?)"&uCodeR
	s=re.Replace(s,tCode)
	re.Pattern=uCodeL
	s=re.Replace(s,"")
	re.Pattern=uCodeR
	s=re.Replace(s,"")
	ubbcode=s
End Function
'--------------------------------------------------------
function ubb(byval s)
	dim ls
	set re=new regExp
	re.IgnoreCase=true
	re.Global=true
	if isnull(s) or s="" then
		ubb=""
		exit function
	end if

	s=html_encode(s)
	ls=lcase(s)
	re.pattern="((javascript:)|(jscript:)|(object)|(js:)|(location.)|(vbscript:)|(vbs:)|(\.value)|(about:)|(file:)|(document.cookie)|(on(mouse|exit|error|click|key|load)))"
	s=re.replace(s,"<i>$1</i>")
	if instr(ls,"[/b]")>0 then s=ubbcode(s,"\[b\]","\[\/b\]","<b>$1</b>")
	if instr(ls,"[/i]")>0 then s=ubbcode(s,"\[i\]","\[\/i\]","<i>$1</i>")
	if instr(ls,"[/u]")>0 then s=ubbcode(s,"\[u\]","\[\/u\]","<u>$1</u>")
	if instr(ls,"[/strike]")>0 then s=ubbcode(s,"\[strike\]","\[\/strike\]","<strike>$1</strike>")
	if instr(ls,"[/align]")>0 then s=ubbcode(s,"\[align=(center|left|right)\]","\[\/align\]","<div style=""text-align:$1"">$2</div>")
	if instr(ls,"[/fly]")>0 then s=ubbcode(s,"\[fly\]","\[\/fly\]","<div style=""text-align:center""><marquee width=50% behavior=alternate scrollamount=2>$1</marquee></div>")
	if instr(ls,"[/move]")>0 then s=ubbcode(s,"\[move\]","\[\/move\]","<marquee scrollamount=2>$1</marquee>")
	if instr(ls,"[/img]")>0 then s=ubbcode(s,"\[img\](http|https|ftp):\/\/","\[\/img\]","<img src=$1://$2 border=0>")
	if instr(ls,"[/sound]")>0 then s=ubbcode(s,"\[sound\]","\[\/sound\]","<a href=""$1"" target=_blank><IMG SRC=images/common/mid.gif border=0 alt='背景音乐'></a><bgsound src=""$1"" loop=""-1"">")
	if instr(ls,"[/color]")>0 then s=ubbcode(s,"\[color=((#.{6})|.{3,6})\]","\[\/color\]","<font color=""$1"">$3</font>")
	if instr(ls,"[/bgcolor]")>0 then s=ubbcode(s,"\[bgcolor=((#.{6})|.{3,6})\]","\[\/bgcolor\]","<font style=""background:$1"">$3</font>")
	if instr(ls,"[/size]")>0 then s=ubbcode(s,"\[size=([1-7])\]","\[\/size\]","<font size=$1>$2</font>")
	if instr(ls,"[/url]")>0 then
		re.pattern="(\[url\])(http:\/\/\S+?)(\[\/url\])"
		s=re.replace(s,"<img src=/images/htm.gif> <a href=""$2"" target='_blank'>$2</a>")
		re.pattern="\[url=(.{5,}?)\](.+?)\[/url\]"
		s=re.replace(s,"<img src=/images/htm.gif> <a href=""$1"" target='_blank'>$2</a>")
	end if

	if instr(ls,"[/iframe]")>0 then s=ubbcode(s,"\[iframe=*([0-9]*),*([0-9]*)\]","\[\/iframe\]","<div style=text-align:center><iframe src=$3 width=$1 height=$2 frameborder=no border=0 marginwidth=0 marginheight=0 scrolling=no></iframe></div>")
	if instr(ls,"[/mp]")>0 then s=ubbcode(s,"\[mp=*([0-9]*),*([0-9]*)\]","\[\/mp\]","<object classid=CLSID:22d6f312-b0f6-11d0-94ab-0080c74c7e95 id=MediaPlayer width=$1 height=$2 ><param name=AUTOSTART VALUE=true ><param name=ShowStatusBar value=-1><param name=Filename value=$3><embed type=application/x-oleobject codebase=http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701 flename=mp src=$3 width=$1 height=$2></embed></object>")
	if instr(ls,"[/rm]")>0 then s=ubbcode(s,"\[RM=*([0-9]*),*([0-9]*)\]","\[\/rm\]","<OBJECT classid=clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA id=RAOCX width=$1 height=$2>" & vbcrlf & "<PARAM NAME=SRC VALUE=$3>" & vbcrlf & "<PARAM NAME=CONSOLE VALUE=Clip1>" & vbcrlf & "<PARAM NAME=CONTROLS VALUE=imagewindow>" & vbcrlf & "<PARAM NAME=AUTOSTART VALUE=false>" & vbcrlf & "</OBJECT>" & vbcrlf & "<br>" & vbcrlf & "<OBJECT classid=CLSID:CFCDAA03-8BE4-11CF-B84B-0020AFBBCCFA height=32 id=video2 width=$1>" & vbcrlf & "<PARAM NAME=SRC VALUE=$3>" & vbcrlf & "<PARAM NAME=AUTOSTART VALUE=-1>" & vbcrlf & "<PARAM NAME=CONTROLS VALUE=controlpanel>" & vbcrlf & "<PARAM NAME=CONSOLE VALUE=Clip1>" & vbcrlf & "</OBJECT>")
	if instr(ls,"[/ra]")>0 then s=ubbcode(s,"\[ra\]","\[\/ra\]","<object classid=clsid:CFCDAA03-8BE4-11CF-B84B-0020AFBBCCFA class=OBJECT id=RAOCX width=280 height=70><param name=_ExtentX value=7938><param name=_ExtentY value=2646><param name=AUTOSTART value=-1><param name=SHUFFLE value=0><param name=PREFETCH value=0><param name=NOLABELS value=0><param name=LOOP value=0><param name=NUMLOOP value=0><param name=CENTER value=0><param name=MAINTAINASPECT value=0><param name=BACKGROUNDCOLOR value=#000000><PARAM NAME=SRC VALUE=$1></object>")
	if instr(ls,"[/flash]")>0 then s=ubbcode(s,"\[flash=*([0-9]*),*([0-9]*)\]","\[\/flash\]","<OBJECT codeBase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0 classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 width=$1 height=$2><PARAM NAME=movie VALUE=""$3""><PARAM NAME=quality VALUE=high><embed src=""$3"" quality=high pluginspage=''http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash'' type=''application/x-shockwave-flash'' width=$1 height=$2>$3</embed></OBJECT>")

	'自动识别网址
	If InStr(ls,"http")>0 Or InStr(ls,"https")>0 Or InStr(ls,"ftp")>0 Or InStr(ls,"rtsp")>0 Or InStr(ls,"mms")>0 Then
		re.Pattern = "(^|[^<=""])(((http|https|ftp|rtsp|mms):(\/\/|\\\\))(([\w\/\\\+\-~`@:%])+\.)+([\w\/\\\.\=\?\+\-~`@\':!%#]|(&amp;)|&)+)"
		s = re.Replace(s,"$1<a target=_blank href=$2>$2</a>")
	End If
	'自动识别www等开头的网址
	If InStr(ls,"www")>0 Then
		re.Pattern = "(^|[^\/\\\w\=])((www|bbs)\.(\w)+\.([\w\/\\\.\=\?\+\-~`@\'!%#]|(&amp;))+)"
		s = re.Replace(s,"$1<a target=_blank href=http://$2>$2</a>")
	End If

	dim tuid
	tuid=cnum(request.cookies("tuserid"))
	totable=cnum(request.cookies("tb"))
	if totable="" or totable<1 then totable=1
	ubb=s
	set re=nothing
end function
%>