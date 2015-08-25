
<%
Public Function Ubbcode(strcontent)
	dim re
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True

	'strcontent=Replace(strcontent,"file:","file :")
	'strcontent=Replace(strcontent,"files:","files :")
	'strcontent=Replace(strcontent,"script:","script :")
	'strcontent=Replace(strcontent,"js:","js :")
    
    '图片UBB
	re.pattern="\[img\](http|https|ftp):\/\/(.[^\[]*)\[\/img\]"
	strcontent=re.replace(strcontent,"<a onfocus=""this.blur()"" href=""$1://$2"" target=new><img src=""$1://$2"" border=""0"" alt=""按此在新窗口浏览图片"" onload=""javascript:if(this.width>screen.width-333)this.width=screen.width-333""></a>")
	'链接UBB
	re.pattern="(\[url\])(.[^\[]*)(\[url\])"
	strcontent= re.replace(strcontent,"<img align=""absmiddle"" src=""images/url1.gif""><a href=""$2"" target=""new"">$2</a>")
	re.pattern="\[url=(.[^\[]*)\]"
	strcontent= re.replace(strcontent,"<img align=""absmiddle"" src=""images/url1.gif""><a href=""$1"" target=""new"">")
	'邮箱UBB
	re.pattern="(\[email\])(.*?)(\[\/email\])"
	strcontent= re.replace(strcontent,"<img align=""absmiddle"" src=""images/email1.gif""><a href=""mailto:$2"">$2</a>")
	re.pattern="\[email=(.[^\[]*)\]"
	strcontent= re.replace(strcontent,"<img align=""absmiddle"" src=""images/email1.gif""><a href=""mailto:$1"" target=""new"">")
	'QQ号码UBB
	're.pattern="\[qq=([1-9]*)\]([1-9]*)\[\/qq\]"
	'strcontent= re.replace(strcontent,"<a target=""new"" href=""tencent://message/?uin=$2&Site=www.52515.net&Menu=yes""><img border=""0"" src=""http://wpa.qq.com/pa?p=1:$2:$1"" alt=""点击这里给我发消息""></a>")
    '颜色UBB
	re.pattern="\[color=(.[^\[]*)\]"
	strcontent=re.replace(strcontent,"<font color=""$1"">")
	'文字字体UBB
	re.pattern="\[font=(.[^\[]*)\]"
	strcontent=re.replace(strcontent,"<font face=""$1"">")
	'文字大小UBB
	re.pattern="\[size=([1-7])\]"
	strcontent=re.replace(strcontent,"<font size=""$1"">")
	'文字对齐方式UBB
	re.pattern="\[align=(center|left|right)\]"
	strcontent=re.replace(strcontent,"<div align=""$1"">")
	'表格UBB
	re.pattern="\[table=(.[^\[]*)\]"
	strcontent=re.replace(strcontent,"<table width=""$1"">")

    'FLASH动画UBB
	re.pattern="(\[flash\])(http://.[^\[]*(.swf))(\[\/flash\])"
	strcontent= re.replace(strcontent,"<a href=""$2"" target=""new""><img src=""images/swf.gif"" border=""0"" alt=""点击开新窗口欣赏该flash动画!"" height=""16"" width=""16"">[全屏欣赏]</a><br><center><object codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0"" classid=""clsid:d27cdb6e-ae6d-11cf-96b8-444553540000"" width=""300"" height=""200""><param name=""movie"" value=""$2""><param name=""quality"" value=""high""><embed src=""$2"" quality=""high"" pluginspage=""http://www.macromedia.com/shockwave/download/index.cgi?p1_prod_version=shockwaveflash"" type=""application/x-shockwave-flash"" width=""300"" height=""200"">$2</embed></object></center>")
    
	re.pattern="(\[flash=*([0-9]*),*([0-9]*)\])(http://.[^\[]*(.swf))(\[\/flash\])"
	strcontent= re.replace(strcontent,"<a href=""$4"" target=""new""><img src=""images/swf.gif"" border=""0"" alt=""点击开新窗口欣赏该flash动画!"" height=""16"" width=""16"">[全屏欣赏]</a><br><center><object codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0"" classid=""clsid:d27cdb6e-ae6d-11cf-96b8-444553540000"" width=""$2"" height=""$3""><param name=""movie"" value=""$4""><param name=quality value=high><embed src=""$4"" quality=""high"" pluginspage=""http://www.macromedia.com/shockwave/download/index.cgi?p1_prod_version=shockwaveflash"" type=""application/x-shockwave-flash"" width=""$2"" height=""$3"">$4</embed></object></center>")

    'MEDIA PLAY播放UBB
	re.pattern="\[wmv\](.[^\[]*)\[\/wmv]"
	strcontent=re.replace(strcontent,"<object align=""middle"" classid=""clsid:22d6f312-b0f6-11d0-94ab-0080c74c7e95"" class=""object"" id=""mediaplayer"" width=""300"" height=""200"" ><param name=""showstatusbar"" value=""-1""><param name=""filename"" value=""$1""><embed type=""application/x-oleobject"" codebase=""http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#version=5,1,52,701"" flename=""mp"" src=""$1""  width=""300"" height=""200""></embed></object>")

	re.pattern="\[wmv=*([0-9]*),*([0-9]*)\](.[^\[]*)\[\/wmv]"
	strcontent=re.replace(strcontent,"<object align=""middle"" classid=""clsid:22d6f312-b0f6-11d0-94ab-0080c74c7e95"" class=""object"" id=""mediaplayer"" width=""$1"" height=""$2"" ><param name=""showstatusbar"" value=""-1""><param name=""filename"" value=""$3""><embed type=""application/x-oleobject"" codebase=""http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#version=5,1,52,701"" flename=""mp"" src=""$3""  width=""$1"" height=""$2""></embed></object>")

    
	'REALPLAY 播放UBB
    re.pattern="\[rm\](.[^\[]*)\[\/rm]"
	strcontent=re.replace(strcontent,"<object classid=""clsid:cfcdaa03-8be4-11cf-b84b-0020afbbccfa"" class=""object"" id=""raocx"" width=""300"" height=""200""><param name=""src"" value=""$1""><param name=""console"" value=""clip1""><param name=""controls"" value=""imagewindow""><param name=""autostart"" value=""true""></object><br><object classid=""clsid:cfcdaa03-8be4-11cf-b84b-0020afbbccfa"" height=""32"" id=""video2"" width=""300""><param name=""src"" value=""$1""><param name=""autostart"" value=""-1""><param name=""controls"" value=""controlpanel""><param name=""console"" value=""clip1""></object>")

	re.pattern="\[rm=*([0-9]*),*([0-9]*)\](.[^\[]*)\[\/rm]"
	strcontent=re.replace(strcontent,"<object classid=""clsid:cfcdaa03-8be4-11cf-b84b-0020afbbccfa"" class=""object"" id=""raocx"" width=""$1"" height=""$2""><param name=""src"" value=""$3""><param name=""console"" value=""clip1""><param name=""controls"" value=""imagewindow""><param name=""autostart"" value=""true""></object><br><object classid=""clsid:cfcdaa03-8be4-11cf-b84b-0020afbbccfa"" height=""32"" id=""video2"" width=""$1""><param name=""src"" value=""$3""><param name=""autostart"" value=""-1""><param name=""controls"" value=""controlpanel""><param name=""console"" value=""clip1""></object>")

    strcontent=replace(strcontent,vbcrlf,"<BR/>")
	re.pattern="\[code\]((.|\n)*?)\[\/code\]"
	Set tempcodes=re.Execute(strcontent)
	For i=0 To tempcodes.count-1
	  re.pattern="<BR/>"
	  tempcode=Replace(tempcodes(i),"<BR/>",vbcrlf)
	  strcontent=replace(strcontent,tempcodes(i),tempcode)
	next

    searcharray=Array("[/url]","[/email]","[/color]", "[/size]", "[/font]", "[/align]", "[b]", "[/b]","[i]", "[/i]", "[u]", "[/u]", "[list]", "[list=1]", "[list=a]","[list=A]", "[*]", "[/list]", "[indent]", "[/indent]","[code]","[/code]","[quote]","[/quote]","[table]","[tr]","[td]","[/tr]","[/td]","[/table]")
	replacearray=Array("</a>","</a>","</font>", "</font>", "</font>", "</div>", "<b>", "</b>", "<i>","</i>", "<u>", "</u>", "<ul>", "<ol type=1>", "<ol type=a>","<ol type=A>", "<li>", "</ul></ol>", "<blockquote>", "</blockquote>","<div><textarea name=""codes"" id=""codes"" rows=""12"" cols=""65"">","</textarea><br/><input type=""button"" value=""运行代码"" onclick=""RunCode()""> <input type=""button"" value=""复制代码"" onclick=""CopyCode()""> <input type=""button"" value=""另存代码"" onclick=""SaveCode()""> &nbsp;提示：您可以先修改部分代码再运行</div>","<div style=""background:#E2F2FF;width:90%;height:300px;border:1px solid #3CAAEC"">","</div>","<table>","<tr>","<td>","</tr>","</td>","</table>")
	For i=0 To UBound(searcharray)
		strcontent=replace(strcontent,searcharray(i),replacearray(i))
	next
	set re=Nothing
	Ubbcode=strcontent
End Function

're.Pattern="\[UPLOAD=(gif|jpg|jpeg|bmp)\](.[^\[]*)(gif|jpg|jpeg|bmp)\[\/UPLOAD\]"
'strcontent= re.Replace(strcontent,"<br><IMG SRC=""images/$1.gif"" border=0>此主题相关链接如下：<br><A HREF=""$2$1"" TARGET=_blank><IMG SRC=""$2$1"" border=0 alt=按此在新窗口浏览图片 onload=""javascript:if(this.width>screen.width-333)this.width=screen.width-333""></A>")
're.Pattern="\[UPLOAD=(doc|xls|ppt|htm|swf|rar|zip|exe)\](.[^\[]*)(doc|xls|ppt|htm|swf|rar|zip|exe)\[\/UPLOAD\]"
'strcontent= re.Replace(strcontent,"<br><IMG SRC=""images/$1.gif"" border=0>此主题相关链接如下：<br><a href=""$2$1"" target='_blank'>点击浏览该文件</a>")

'自动识别网址
're.Pattern = "^((http|https|ftp|rtsp|mms):(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@[\]\':+!]+)"
'strcontent = re.Replace(strcontent,"<img align=absmiddle src=images/url.gif border=0><a target=_blank href=$1>$1</a>")
're.Pattern = "((http|https|ftp|rtsp|mms):(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@[\]\':+!]+)$"
'strcontent = re.Replace(strcontent,"<img align=absmiddle src=images/url.gif border=0><a target=_blank href=$1>$1</a>")
're.Pattern = "([^>=""])((http|https|ftp|rtsp|mms):(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@[\]\':+!]+)"
'strcontent = re.Replace(strcontent,"$1<img align=absmiddle src=images/url.gif border=0><a target=_blank href=$2>$2</a>")

'自动识别www等开头的网址
're.Pattern = "([^(http://|http:\\)])((www|cn)[.](\w)+[.]{1,}(net|com|cn|org|cc)(((\/[\~]*|\\[\~]*)(\w)+)|[.](\w)+)*(((([?](\w)+){1}[=]*))*((\w)+){1}([\&](\w)+[\=](\w)+)*)*)"
'strcontent = re.Replace(strcontent,"<img align=absmiddle src=images/url.gif border=0><a target=_blank href=http://$2>$2</a>")

're.Pattern="\[SHADOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\](.[^\[]*)\[\/SHADOW]"
'strcontent=re.Replace(strcontent,"<div style=""width:$1;filter:shadow(color=$2, strength=$3)"">$4</div>")
're.Pattern="\[GLOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\](.[^\[]*)\[\/GLOW]"
'strcontent=re.Replace(strcontent,"<div style=""width:$1;filter:glow(color=$2, strength=$3)"">$4</div>")
%>