<%@ CODEPAGE = "936" %>
<%
'变量定义后可用
Option Explicit
Response.CodePage=936
Response.Charset="gb2312" 
Session.CodePage=936

'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True 
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 

Dim PageStartTime
PageStartTime=Timer()
'=======================================================================
'开始数据库连接
Dim ConnStr,Conn,Rs,Data_Path
'数据库文件路径
'Data_Path="../data/data.asp"
Data_Path="databases/data.mdb"
If IsFile(Data_Path)=False Then Data_Path="../" & Data_Path
ConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(Data_Path)
'========================================================================
Dim SystemKey,User_ID,User_Group,User_Pass,User_VIP,User_Ip,Copyright,En_Only,Sql,Num,i,Key,Act,boardCount,configChg,Adm_UP_FileSize,Adm_UP_FileType,Mem_UP_FileSize,Mem_UP_FileType,vip_UP_FileType,vip_Up_FileSize,MemCanUP,allowNum,css_style,sMode,strUrl,strUrl2,strUrl2s,strSiteName,strCopyright,strSiteUrl,strSiteLogo,strAbout,strBeiAn,strKeyword,contextPagesize,ListPagesize,blogPagesize,defaultPost,actstr,logontype,strContentsize,strStyle,indexPost,css_name,readmode,replymode,postmode,strBadWords,strSpamIp,strEditor,strBanner,forumid,strtongji,strfooterin,Hotkw,strAdminEmail
Dim strRegTsEmail,strSubTsEmail,strReTsEmail,strPostTsEmail,strReSsbTsEmail
Dim Enable_Register,Enable_Anonymous,Enable_Ubb,Enable_Cookies,Enable_Upload,Enable_UBB_Media,Enable_SMS,Enable_Emot,EnableClass,EnableRead,EnablePost,EnableReply,Enable_VLCode
SystemKey="mcms" 
'========================================================================
'读入基本配置
%>
<!--#include file="config.asp"-->
<%
'========================================================================
'去除查询漏洞
Function SqlShow(Str)
	SqlShow=Replace(Str,"'","''")
End Function
'========================================================================
'打开数据库连接
Sub OpenData()
	Set Conn=Server.CreateObject("Adodb.Connection")
	Set Rs=Server.CreateObject("Adodb.Recordset")
	Conn.Open ConnStr
    If Err Then
       err.Clear
       Set conn = Nothing
        response.write("数据库连接出错，请修正文件")
        Response.End
	End If
End Sub
'========================================================================
'关闭数据库连接
Sub CloseData()
	Set Rs=Nothing
	Set Conn=Nothing
End Sub
'========================================================================
'关闭数据库及数据集
Sub CloseAll()
	Set Rs=Nothing
	Conn.Close
	Set Conn=Nothing
End Sub
'========================================================================
'打开数据集
Sub OpenRs(ByVal SqlStr)
	If Left(LCase(SqlStr),6)="select" Then
		Rs.Open SqlStr,Conn,1,3
	Else
		Conn.Execute SqlStr
	End If
End Sub
'========================================================================
'检查用户登录状态
Sub CheckUser()
If Enable_Cookies = 1 Then
  If Session(SystemKey & "User_ID")="" Then
	dim userid,userpassword,usergroup,uservip
	userid=Request.cookies(systemkey)("userid")
	userpassword=Request.cookies(systemkey)("userpassword")
	usergroup=Request.cookies(systemkey)("usergroup")
	uservip=Cnum(Request.cookies(systemkey)("uservip"))
	If userid<>"" Then
		sql="select * from users where username='"&userid&"'and userpass='"&userpassword&"'and usergroup='"&usergroup&"' and vip="&uservip&""
		openrs(sql)
		if  rs.EOF and rs.BOF Then
        Response.cookies(systemkey)("userid")=""
		Response.cookies(systemkey)("usergroup")=""
		Response.cookies(systemkey)("userpassword")=""
		Session(SystemKey & "User_ID")=""
		Session(SystemKey & "User_Group")=""
		Session(SystemKey&"User_VIP")=""
		else
		If Rs("lockuser")=TRUE Then
		Call ShowError("用户已被管理员锁定，不能登录！",1)
		Else
		Session(SystemKey & "User_ID")=rs("username")
        Session(SystemKey & "User_Group")=rs("usergroup")
		Session(SystemKey&"User_VIP")=rs("vip")
		End If
		end if
		Rs.Close
	Else
		Response.cookies(systemkey)("userid")=""
		Response.cookies(systemkey)("usergroup")=""
		Response.cookies(systemkey)("userpassword")=""
		Session(SystemKey & "User_ID")=""
		Session(SystemKey & "User_Group")=""
		Session(SystemKey&"User_VIP")=""
	End If
  End If
	User_ID=Session(SystemKey & "User_ID")
	User_Group=Session(SystemKey & "User_Group")
	User_VIP=Session(SystemKey & "User_VIP")
Else
	User_ID=Session(SystemKey & "User_ID")
	User_Group=Session(SystemKey & "User_Group")
	User_VIP=Session(SystemKey & "User_VIP")
End If
End Sub
'========================================================================
'把一个字符变成一个数
Function Cnum(Num)
	If IsNumeric(Num) Then
		Cnum=Clng(Num)
	Else
		Cnum=0
	End If
End Function
'========================================================================
'求两数大者
Function Max(Num1,Num2)
	If Num1>Num2 Then
		Max=Num1
	Else
		Max=Num2
	End If
End Function
'========================================================================
'求两数小者
Function Min(Num1,Num2)
	If Num1>Num2 Then
		Min=Num2
	Else
		Min=Num1
	End If
End Function
'========================================================================
'判断文件是否存在
Function IsFile(tPath)
	Dim Fso,Path
	Set Fso=CreateObject("Scripting.FileSystemObject")
	If Mid(tPath,2,1)=":" Then
		Path=tPath
	Else
		Path=Server.MapPath(tPath)
	End If
	IsFile=Fso.FileExists(Path)
	Set Fso=Nothing
End Function
'========================================================================
'仅返回日期中的年月日
Function OnlyDate(ExpStr)
	Dim DateYear,DateMonth,DateDay
	DateYear=Year(ExpStr)
	DateMonth=Month(ExpStr)
	DateDay=Day(ExpStr)
	If Len(DateMonth)<2 Then DateMonth="0"&DateMonth
	If Len(DateDay)<2 Then DateDay="0"&DateDay
	OnlyDate=DateYear&"-"&DateMonth&"-"&DateDay
End Function
'========================================================================
'仅返回时间
Function OnlyTime(ExpStr)
	Dim DateHour,DateMinute,DateSecond
	DateHour=Hour(ExpStr)
	DateMinute=Minute(ExpStr)
	DateSecond=Second(ExpStr)
	If Len(DateHour)<2 Then DateHour="0"&DateHour	
	If Len(DateMinute)<2 Then DateMinute="0"&DateMinute
	If Len(DateSecond)<2 Then DateSecond="0"&DateSecond
	OnlyTime=DateHour&":"&DateMinute&":"&DateSecond
End Function
'========================================================================
'分面显示
Function ShowPages(Pages,Page,Url)	
Dim i,Str,FrontStr,BackStr,ShowStr,StartNum,EndNum
Str=Url
If Replace(Str,"?","")<>Str Then
	Str=Str & "&page="
Else
	Str=Str & "?page="
End If
FrontStr="<a href=""" & Str & 1 & """ title=""第一页""><</a>"
BackStr="<a href=""" & Str & Pages & """ title=""最后一页"">></a>"
If Pages<=1 Then
	ShowPages=""
	Exit Function
End If
If Pages<=8 Then
	For i=1 To Pages
		If i<>Page Then
			ShowPages=ShowPages & "<a href="""& Str & i & """>" & i & "</a> "
		Else
			ShowPages=ShowPages & "<a class=""tebo"">" & i & "</a>"
		End If
	Next
	ShowPages=FrontStr & " " &  ShowPages & " " & BackStr
	Exit Function
End If
If Pages>8 Then
	StartNum=Page-5
	EndNum=StartNum+7
	If StartNum<=0 Then
		StartNum=1
		EndNum=StartNum+7
	End If
	If EndNum>Pages Then
		EndNum=Pages
		StartNum=EndNum-7
	End If
	For i=StartNum To EndNum
		If i<>Page Then
			If i=Pages Then
				ShowPages=ShowPages & "<a href=""" & Str & Pages & """>" & Pages & "</a>"
				ShowPages=ShowPages & "<a href=""" & Str & Pages & """ title=""最后一页"">></a>"
			Else				
				ShowPages=ShowPages & "<a href=""" & Str & i & """>" & i & "</a> "
			End If
		Else
			If i=Pages Then
				ShowPages=ShowPages & "<a class=""tebo"" title=""最后一页"">" & Pages & "</a>"
			Else
				ShowPages=ShowPages & "<a class=""tebo"">" & i & "</a> "
			End If
		End If
	Next
	ShowPages=FrontStr & " " & ShowPages
	If EndNum<Pages Then
		ShowPages=ShowPages & "...&nbsp;<a href=""" & Str & Pages & """ title=""最后一页"">" & Pages & "</a><a href=""" & Str & Pages & """ title=""最后一页"">></a>"
	End If
End If
End Function
'========================================================================
'取对方IP地址
Function RemoteIp()	
	If Request.ServerVariables("HTTP_X_FORWARDED_FOR")<>"" then 
		RemoteIp=Request.ServerVariables("HTTP_X_FORWARDED_FOR")
	Else
		RemoteIp=Request.ServerVariables("REMOTE_ADDR")
	End If
End Function
'========================================================================
'页面处理时长
Function TillNowTime()
	Dim NowTime
	NowTime=Timer()
	TillNowTime=Round((NowTime-PageStartTime)*1000,4)
End Function
'========================================================================
'页重定向
Function TurnTo(ByVal URl)
	On Error Resume Next
	Rs.Close
	CloseAll
	Response.Clear
	Response.Redirect(URL)
End Function
'=========================================================================
'字符串长度，支持中英文混用
function strlen(str)
dim p_len,xx 
p_len=0 
strlen=0 
if trim(str)<>"" then 
p_len=len(trim(str)) 
for xx=1 to p_len 
if asc(mid(str,xx,1))<0 then 
strlen=int(strlen) + 2 
else 
strlen=int(strlen) + 1 
end if 
next 
end if 
end function 
'================================================================
'截取字符串，支持中英文混用
function strvalue(str,lennum)
dim p_num,x 
dim i 
if strlen(str)<=lennum then 
strvalue=str 
else 
p_num=0 
x=0 
do while not p_num > lennum-2 
x=x+1 
if asc(mid(str,x,1))<0 then 
p_num=int(p_num) + 2 
else 
p_num=int(p_num) + 1 
end if 
strvalue=left(trim(str),x)&"…" 
loop 
end if 
end function 
'=========================================================================
'截取字符串，支持中英文混用
Function CutStr(byVal Str,byVal StrLen) 
     Dim l,t,c,i 
     l=Len(str) 
     t=0 
     For i=1 To l 
           c=AscW(Mid(str,i,1)) 
           If c<0 Or c>255 Then t=t+2 Else t=t+1 
           IF t>=StrLen Then 
                 CutStr=left(Str,i)&"..." 
                 Exit For 
           Else 
                 CutStr=Str 
           End If 
     Next 
End Function 
'=========================================================================
'文章中的纯文字
Function delHtml(strHtml) '做了一个函数名叫delhtml
Dim objRegExp, strOutput
Set objRegExp = New Regexp ' 建立正则表达式
objRegExp.IgnoreCase = True ' 设置是否区分大小写
objRegExp.Global = True '是匹配所有字符串还是只是第一个
objRegExp.Pattern = "([[a-zA-Z].*?])|([[\/][a-zA-Z].*?])" ' 设置模式引号中的是正则表达式，用来找出html标签
strOutput = objRegExp.Replace(strHtml, "") '将html标签去掉
strOutput = Replace(strOutput, "[", "[") '防止非html标签不显示
strOutput = Replace(strOutput, "]", "]") 
delHtml = strOutput
Set objRegExp = Nothing
End Function
'=========================================================================
'转换HTML代码
Function HTMLEncode(reString)
	Dim Str:Str=reString
	If Not IsNull(Str) Then
		Str = UnCheckStr(Str)
		Str = Replace(Str, "&", "&amp;")
		Str = Replace(Str, ">", "&gt;")
		Str = Replace(Str, "<", "&lt;")
		Str = Replace(Str, CHR(34),"&quot;")
		Str = Replace(Str, CHR(39),"&#39;")
		Str = Replace(Str, CHR(13), "")
		Str = Replace(Str, "       ", "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;", 1, -1, 0)
		Str = Replace(Str, "      ", "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;", 1, -1, 0)
		Str = Replace(Str, "     ", "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;", 1, -1, 0)
		Str = Replace(Str, "    ", "&nbsp;&nbsp;&nbsp;&nbsp;", 1, -1, 0)
		Str = Replace(Str, "   ", "&nbsp;&nbsp;&nbsp;", 1, -1, 0)
		Str = Replace(Str, "  ", "&nbsp;&nbsp;", 1, -1, 0)
		Str = Replace(Str, CHR(10), "<br>")
		Str	= chkBadWords(Str)
		HTMLEncode = Str
	End If
End Function
'==========================================================================
Function UnCheckStr(Str)
		Str = Replace(Str, "sel&#101;ct", "select")
		Str = Replace(Str, "jo&#105;n", "join")
		Str = Replace(Str, "un&#105;on", "union")
		Str = Replace(Str, "wh&#101;re", "where")
		Str = Replace(Str, "ins&#101;rt", "insert")
		Str = Replace(Str, "del&#101;te", "delete")
		Str = Replace(Str, "up&#100;ate", "update")
		Str = Replace(Str, "lik&#101;", "like")
		Str = Replace(Str, "dro&#112;", "drop")
		Str = Replace(Str, "cr&#101;ate", "create")
		Str = Replace(Str, "mod&#105;fy", "modify")
		Str = Replace(Str, "ren&#097;me", "rename")
		Str = Replace(Str, "alt&#101;r", "alter")
		Str = Replace(Str, "ca&#115;t", "cast")
		UnCheckStr=Str
End Function
'==================================================================
Function formatIP(IpAddress)
	dim IpArray,IpAdd,x
	IpArray=split(IpAddress,".",-1,1)
	For x = Lbound(IpArray) to Ubound(IpArray)-1 
		IpAdd=IpAdd+IpArray(x)+"."
	Next 
	formatIP=IpAdd+"*"
End Function
'==================================================================
Function formatIP2(IpAddress)
	dim IpArray,IpAdd,x
	IpArray=split(IpAddress,".",-1,1)
	For x = Lbound(IpArray) to Ubound(IpArray)-2 
		IpAdd=IpAdd+IpArray(x)+"."
	Next 
	formatIP2=IpAdd+"*.*"
End Function
'===========================================================================
'友情链接
Sub friendsite()
dim Rs2,sql2
sql2="select * from friendsite order by siteorder"
Set Rs2=Server.CreateObject("Adodb.Recordset")
Rs2.Open sql2,Conn,1,3
If rs2.recordcount>0 Then
for i = 1 to rs2.recordcount
Response.write ("<li><a target=_blank href="""&Rs2("siteurl")&"""  title="&Rs2("sitename")&">"&Rs2("sitename")&"</a></li>")
Rs2.movenext
next
Else
Response.write("暂无友情链接!")
End If
rs2.close
set rs2=nothing
End Sub
'===========================================================================
'生成公告.txt文件
Sub NoticeSite()
dim Rs2,sql2,sitenum,tmpStr
Sql="Select * From Notice WHERE EnableNotice=True order by OrderID DESC"
Openrs(sql)
For i = 1 to Rs.Recordcount
tmpStr=tmpStr&"marqueeContent["&i-1&"]='<a href="""&rs("Nurl")&""" target=""_blank"">"&rs("NTitle")&"</a>';"&vbcrlf&""
Rs.Movenext
Next
Rs.Close
Set Rs=nothing
Call Save("../txt/Notice.txt",tmpStr)
End Sub
'========================================================================
'读入公告
Sub Notice()
%>
<script> 
var marqueeContent=new Array();   //滚动主题
<!--#include file="../txt/Notice.txt"-->
var marqueeInterval=new Array();  //定义一些常用而且要经常用到的变量
var marqueeId=0;
var marqueeDelay=4000;
var marqueeHeight=24;
function initMarquee() {
 var str=marqueeContent[0];
 document.write('<div id=marqueeBox style="overflow:hidden;height:'+marqueeHeight+'px" onmouseover="clearInterval(marqueeInterval[0])" onmouseout="marqueeInterval[0]=setInterval(\'startMarquee()\',marqueeDelay)"><div>'+str+'</div></div>');
 marqueeId++;
 marqueeInterval[0]=setInterval("startMarquee()",marqueeDelay);
 }
function startMarquee() {
 var str=marqueeContent[marqueeId];
  marqueeId++;
 if(marqueeId>=marqueeContent.length) marqueeId=0;
 if(marqueeBox.childNodes.length==1) {
  var nextLine=document.createElement('DIV');
  nextLine.innerHTML=str;
  marqueeBox.appendChild(nextLine);
  }
 else {
  marqueeBox.childNodes[0].innerHTML=str;
  marqueeBox.appendChild(marqueeBox.childNodes[0]);
  marqueeBox.scrollTop=0;
  }
 clearInterval(marqueeInterval[1]);
 marqueeInterval[1]=setInterval("scrollMarquee()",10);
 }
function scrollMarquee() {
 marqueeBox.scrollTop++;
 if(marqueeBox.scrollTop%marqueeHeight==marqueeHeight){
  clearInterval(marqueeInterval[1]);
  }
 }
initMarquee();
</script>
<%
End Sub
'========================================================================
'搜索热词
Sub HotKeywords()
%>
<!--#include file="../txt/hotkw.txt"-->
<%
End Sub
'========================================================================
'检查访问状态
Sub CheckVisit()
	Dim DateTime,NowCount,today_count,all_count
	DateTime=Date
	If Request.Cookies(SystemKey & "DateVisitTime")<>"" Then
		Exit Sub
	End If
	Response.Cookies(SystemKey & "DateVisitTime")=Now
    Sql="select * from visitinfo where DateTime=#" & DateTime & "#"
	Openrs(Sql)
	If Rs.RecordCount=0 Then
		Rs.AddNew
		Rs("DateTime")=DateTime
	End If
	Rs("Count")=Rs("Count")+1
	NowCount=Rs("Count")
	Rs.Update
	today_count=Rs("Count")
	Rs.Close
	Sql="select sum(count) from visitinfo"
	OpenRs(Sql)
	all_count=Rs(0)
	Rs.Close
	Application.Lock
	Application(systemkey&"tcount")=today_count
	Application(systemkey&"acount")=all_count
	Application.Unlock
End Sub
'=============================================
'显示上下篇
Function getPreNxt(contextID,forumid,listpage)
dim sql2,rs2,strlength,BoardType

strLength=36
Sql2="select top 1 content.id,content.ClassId,content.title from Content where id>"&contextID&" and classid="&forumid&" and parent=0 order by LastUpdateTime"
set rs2 = Conn.Execute(sql2)

If rs2.eof or rs2.bof Then
Response.write("上一篇：<span>-</span>")
Else
BoardType=rs2("ClassId")
  If BoardType=2 Then
    Response.write("上一篇：<span><a href=""contextsingle.asp?id="&Rs2("id")&""" title="""&HTMLEncode(CutStr(Rs2("title"),strlength))&""">"&HTMLEncode(CutStr(Rs2("title"),strlength))&"</a></span>")
  else
    Response.write("上一篇：<span><a href=""context.asp?id="&Rs2("id")&""" title="""&HTMLEncode(CutStr(Rs2("title"),strlength))&""">"&HTMLEncode(CutStr(Rs2("title"),strlength))&"</a></span>")
  end if
End If
Rs2.close
Sql2="select top 1 content.id,content.ClassId,content.title from Content where id<" & contextID & " and classid="&forumid&" and parent=0 order by LastUpdateTime desc"
set rs2 = Conn.Execute(sql2)

If rs2.eof or rs2.bof Then
Response.write("下一篇：<span>-</span>")
Else
BoardType=rs2("ClassId")
  If BoardType=2 Then
    Response.write("下一篇：<span><a href=""contextsingle.asp?id="&Rs2("id")&""" title="""&HTMLEncode(CutStr(Rs2("title"),strlength))&""">"&HTMLEncode(CutStr(Rs2("title"),strlength))&"</a></span>")
  else
    Response.write("下一篇：<span><a href=""context.asp?id="&Rs2("id")&""" title="""&HTMLEncode(CutStr(Rs2("title"),strlength))&""">"&HTMLEncode(CutStr(Rs2("title"),strlength))&"</a></span>")
  end if
End If
Rs2.close
Set rs2=nothing
End Function
'===================================================================
Sub NewsRelated '相关内容
Dim sql2,Rs2,BoardType
sql2="Select Top 6 * From content WHERE id<>"&ContextID&" and Title like '%"&tags&"%' and ClassId="&Cnum(forumid)&" and SubClass="&Cnum(subclassid)&""
Set Rs2=Server.CreateObject("Adodb.Recordset")
Rs2.Open sql2,Conn,1,3
If rs2.recordcount>0 Then
	For i = 1 To rs2.recordcount
	   BoardType=Rs2("ClassId")
	   If BoardType=2 Then
		   Response.write("<li><h3><a href=""contextsingle.asp?id="&Rs2("ID")&""">"&Rs2("Title")&"</a></h3></li>")
	   else
		   Response.write("<li><h3><a href=""context.asp?id="&Rs2("ID")&""">"&Rs2("Title")&"</a></h3></li>")  
	   end if
	Rs2.movenext
	Next
Else
    Response.write("<li>暂无内容！</li>")
End If
rs2.close
Set rs2=nothing
End Sub
'=============================================
'显示导航名称
Sub getMainNavBar()
If User_ID="" Then 
Response.write("<form action=""include/login.asp?act=login"" method=""post"" name=""form2"">用户名<input type=""text"" name=""username"" size=""8"" tabindex=""1"" class=""dlinput"" />&nbsp;&nbsp;密码<input type=""password"" name=""userpass"" size=""8"" tabindex=""2"" class=""dlinput"" />&nbsp;&nbsp;<input type=""submit"" value=""登录"" name=""B1"" tabindex=""3"" class=""dlbutton"" />&nbsp;&nbsp;<a href=""javascript:openreg();"">注册会员</a></form>") 	
Else
If User_VIP=1 Then Response.write("<span class=""red"">vip*</span>")
Response.write("[<span class=""blue"">"&User_ID&"</span>] ")
'call sms()
Response.write("<a href=""javascript:memberinfo('"&User_ID&"')"">我的资料</a> | <a href=""javascript:openreg();"" class=""gray"">修改</a>")
If User_Group="admin" Then Response.write(" - <a href=""admin/admin_main.asp"">网站管理</a>")
Response.write(" - <a href=""include/login.asp?act=logout"">退出</a>")
End If
End Sub
'=============================================
'显示文章系统栏目名称 双行
Sub getNavBar()
dim tempClassname
If cnum(forumID)=0 Then tempClassname="home" Else tempClassname="home_on"
Response.write("<a href=""" & strSiteUrl & """ class="""&tempClassname&""">首 页</a>")
For j=1 To BoardCount
if Cnum(forumID)=j Then tempClassname="nav_on" Else tempClassname=""
Response.write "<span></span><a class="""&tempClassname&""" href="""&strUrl&"?forumID="&j&""">"&Application(systemkey&j)&"</a>"
Next
End Sub
'=============================================
'显示分类名称
Sub getSubClass()
If forumID<>"" Then
for i=1 to ubound(subclassArray)
If i=cnum(subclassID) Then 
Response.write("<li class=""ClassNameTitle1""><a href="&strUrl&"?forumID="&forumID&"&subclassID="&i&"><font color=""#ba2636"">"&subclassArray(i)&"</font></a></li>")
Else
response.write("<li class=""ClassNameTitle""><a href="&strUrl&"?forumID="&forumID&"&subclassID="&i&">"&subclassArray(i)&"</a></li>")
End If
next
End If
End Sub
'=============================================
'首页显示所有分类名称   ---朱平修改
Sub getAllSubClass()
If forumID<>"" Then
for i=1 to ubound(subclassArray)
If i=cnum(subclassID) Then 
Response.write("<li class=""ClassNameTitle1""><a href="&strUrl&"?forumID="&forumID&"&subclassID="&i&"><font color=""#ba2636"">"&subclassArray(i)&"</font></a></li>")
Else
response.write("<li class=""ClassNameTitle""><a href="&strUrl&"?forumID="&forumID&"&subclassID="&i&">"&subclassArray(i)&"</a></li>")
End If
next
End If
End Sub
'========================================================================
'显示错误信息
Sub ShowError(msg,mode)
Response.write "<script language='javascript'>alert('"&msg&"');"
If mode=0 Then 
Response.write("window.close();</script>") 
Response.end
Else 
Response.write("history.go(-"&mode&");</script>")
Response.end
End If
End Sub
'==========================================================================
'日期转换函数
Function DateToStr(DateTime,ShowType) 
	Dim DateMonth,DateDay,DateHour,DateMinute
	DateMonth=Month(DateTime)
	DateDay=Day(DateTime)
	DateHour=Hour(DateTime)
	DateMinute=Minute(DateTime)
	If Len(DateMonth)<2 Then DateMonth="0"&DateMonth
	If Len(DateDay)<2 Then DateDay="0"&DateDay
	Select Case ShowType
	Case "Y-m-d"  
		DateToStr=Year(DateTime)&"-"&DateMonth&"-"&DateDay
	Case "Y-m-d H:I A"
		Dim DateAMPM
		If DateHour>12 Then 
			DateHour=DateHour-12
			DateAMPM="PM"
		Else
			DateHour=DateHour
			DateAMPM="AM"
		End If
		If Len(DateHour)<2 Then DateHour="0"&DateHour	
		If Len(DateMinute)<2 Then DateMinute="0"&DateMinute
		DateToStr=Year(DateTime)&"-"&DateMonth&"-"&DateDay&" "&DateHour&":"&DateMinute&" "&DateAMPM
	Case "Y-m-d H:I:S"
		Dim DateSecond
		DateSecond=Second(DateTime)
		If Len(DateHour)<2 Then DateHour="0"&DateHour	
		If Len(DateMinute)<2 Then DateMinute="0"&DateMinute
		If Len(DateSecond)<2 Then DateSecond="0"&DateSecond
		DateToStr=Year(DateTime)&"-"&DateMonth&"-"&DateDay&" "&DateHour&":"&DateMinute&":"&DateSecond
	Case "YmdHIS"
		DateSecond=Second(DateTime)
		If Len(DateHour)<2 Then DateHour="0"&DateHour	
		If Len(DateMinute)<2 Then DateMinute="0"&DateMinute
		If Len(DateSecond)<2 Then DateSecond="0"&DateSecond
		DateToStr=Year(DateTime)&DateMonth&DateDay&DateHour&DateMinute&DateSecond	
	Case "ym"
		DateToStr=Right(Year(DateTime),2)&DateMonth
	Case "ymd"
		DateToStr=Right(Year(DateTime),2)&DateMonth&DateDay
	Case "d"
		DateToStr=DateDay
	Case Else
		If Len(DateHour)<2 Then DateHour="0"&DateHour
		If Len(DateMinute)<2 Then DateMinute="0"&DateMinute
		DateToStr=Year(DateTime)&"-"&DateMonth&"-"&DateDay&" "&DateHour&":"&DateMinute
	End Select
End Function
'========================================================================
' 过滤不良词语的功能
Function ChkBadWords(fString)
	Dim strBadWords2,n
	If strBadWords<>"" Then
      strBadWords2 = split(strBadWords, "|")   'BadWords   是数据库中定义的过滤词语
        for n = 0 to ubound(strBadWords2)
            fString = Replace(fString, strBadWords2(n), string(len(strBadWords2(n)),"*")) 
        next
   End if 
    ChkBadWords = fString
End Function
'========================================================================
'缓存变量
OpenData
If Application(SystemKey&"Loaded")=""  Or configChg="YES" Then		'调入栏目名称
Application.Contents.RemoveAll()
Application.Lock

'缓存栏目
Sql="select * from Board where BoardID>0 order by BoardId"		
Openrs(Sql)
Application(systemkey&"boardcount")=Rs.recordcount
If Not Rs.Eof and Not Rs.Bof Then
For i=1 to Rs.Recordcount-1
Application(systemkey&i)=Rs("boardname")

dim Tubb,Tupload
Tubb=1
Tupload=1

Rs.Movenext
Next
Rs.MoveLast
Application(systemkey&999)=Rs("boardname")
Tubb=1
Tupload=1
End If
Rs.Close

'缓存所有子栏目列表，用||分开栏目，用|分开子栏目
Sql="select * from subclass order by subparent,subnum"
Openrs(Sql)
If Not Rs.Eof and Not Rs.Bof Then
For i=1 to Application(systemkey&"boardcount")-1
Application(SystemKey&"subclass"&i)="分类"
Next
Application(SystemKey&"subclass999")="分类"
For i=1 to Rs.RecordCount
Application(SystemKey&"subclass"&Rs("subparent"))=Application(SystemKey&"subclass"&Rs("subparent"))&"|"&Rs("subname")
Rs.MoveNext
Next
End If
Rs.Close

Application(SystemKey&"subclassAll")=Application(Systemkey&"subclass1")
For i=2 to Application(systemkey&"boardcount")
Application(SystemKey&"subclassAll")=Application(SystemKey&"subclassAll")&","&Application(Systemkey&"subclass"&i)
Next
Application(SystemKey&"subclassAll")=Application(SystemKey&"subclassAll")&Application(Systemkey&"subclass999")

'读入表情缓存
If Enable_Emot = 1 Then
Call getEmot()
End If

Application(SystemKey&"Loaded")="YES"
Application.Unlock

End If
'========================================================================
'设定变量值
boardCount=Cnum(application(systemkey&"boardcount"))-1
actstr="login"
Copyright="Powered by <a target=_blank href='http://www.mihao.net'>Mcms3.0_access_GB</a>"
strUrl="List.asp"
strUrl2="context.asp"
strUrl2s="contextsingle.asp"
css_style=Request.cookies(systemkey&"css")("css_style")   '设置风格           '
if css_style="" then css_style=strStyle else css_style=css_style&".css"	
css_name=left(css_style,inStr(css_style,".")-1)
Call CheckUser()							
If User_ID<>"" Then %>
<%
End If
'=================================================================
'★调用代码说明
'dy   "类别",      "栏目号",       分类号,       数量,     标题长
'dy   "mytopic" ,"1"                ,1            ,10        ,20
''参数说明：
'类别取值说明：
'newpost		新发主题 
'newreply		新回主题 
'hotr			热贴,即按点击量排行
'hotc			人气贴,即按回复量排行
'recommend		推荐贴或精华贴
'栏目号         数字         调用栏目ID，空白调用所有栏目（调多个栏目，请用逗号“,”或竖线“|”分隔，如1,2,3或1|2|3）
'分类号 		数字		 调用子栏目ID
'数量           数字         调用数量，空白则调用10个
'标题长 		数字         主题截取长度
'=================================================================
Sub dy(act,forumid,subclassid,number,t_length)
dim wherestr,orderstr,title,showdate,dyurl,dylm,dydao,neworold,neworcommon,poster,tempClass,subclassArray,BoardType
subclassid=cnum(subclassid)

'初始化参数默认值
if number="" or number=0 then number=12
if t_length="" or t_length=0 then t_length=40
if dyurl="" then dyurl="./"
if instr(right(dyurl,1),"/")=0  then  dyurl=dyurl&"/" 
t_length=Cnum(t_length)

'读取子栏目
If Cnum(forumid)>0 Then
tempClass=Application(SystemKey&"subclass"&forumid)
If instr(tempClass,"|")>0 Then
subclassArray=split(tempClass,"|")
Else
subclassArray=split("分类","|")
End If
End If

If forumid="" Then 
	wherestr="parent=0 and classID<998"
Else
	wherestr="parent=0 and classID in("&forumid&")"
	if subclassID>0 then
		wherestr=wherestr&" and subclass="&subclassID&""
	end if
End if

Select Case Act
	Case "recommend" 
		wherestr="classic=True and  "&wherestr
		orderstr="order by PostTime desc,id desc"
	Case "newpost"
		orderstr="order by posttime desc,id desc"
	Case "newreply"
		wherestr="posttime<>lastupdatetime"
		orderstr="order by lastupdatetime desc,id desc"
	Case "hotc" 
		orderstr="order by clicknum desc,id desc"
	Case "hotr" 
		orderstr="order by replynum DESC,id DESC"
	Case Else
		orderstr="order by posttime DESC,id DESC"
End Select

Sql="Select Top "&Cnum(number)&" Content.Id,Content.Title,Content.ClassID,Content.Classic,Content.TitleColor,Content.TitleBold,Content.subclass From content WHERE "&wherestr&" "&orderstr&""
Openrs(sql)

If Rs.eof and Rs.bof Then
	Response.Write "<center>该栏目下没有内容！</center>"
	Rs.Close
	Exit Sub
End If

For i = 1 to Rs.Recordcount
if strlen(poster)>10 Then poster=CutStr(poster,9)
title=Rs("Title")
If strlen(title)>t_length Then title=CutStr(title,t_length)
If Rs("TitleBold") = True Then title="<b>"&title&"</b>"
If Rs("TitleColor")<>"" Then title="<font color="""&Rs("TitleColor")&""">"&title&"</font>"

If act<>"" Then
     BoardType=Rs("ClassId")
	 If BoardType=2 Then
		 Response.write "<li><a href="""&strUrl2s&"?id="&Rs("ID")&""">"&title&"</a></li>"
	 else
		 Response.write "<li><a href="""&strUrl2&"?id="&Rs("ID")&""">"&title&"</a></li>"
	 end if
    

End If

Rs.Movenext
Next
Rs.Close
End Sub

dim tempClass,subclassArray,subclassID,tempForum
tempForum=ForumID
If ForumID="" Then tempForum=DefaultPost
tempClass=Application(SystemKey&"subclass"&tempForum)
If instr(tempClass,"|")>0 Then
subclassArray=split(tempClass,"|")
Else
subclassArray=split("分类","|")
End If
%>
