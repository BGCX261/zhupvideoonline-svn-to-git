<%
'-------------------------------------------------------------------
'过程名与变量：dyfPic(f_width,f_height,t_height,act,forumid,number,TitleLen,picadd)
'★调用举例：  call dyfPic(300,208,19,"new","all",6,15,"att") 
'           
'★参数说明：
'参数名称     取值          说明
'f_width     数字        设置FLASH幻灯图片宽度
'f_height    数字        设置FLASH幻灯图片高度
't_height    数字        设置FLASH幻灯文字标题高度
'act        (●见下）    调用新贴/调用热贴
'forumid    (见右说明）  调用栏目ID，"all"或""空白调用所有栏目（调多个栏目，请用逗号“,”或竖线“|”分隔，如1,2,3或1|2|3）
'number      数字        调用数量,最多6个，空白或0则调用6个(不超过6个）
'TitleLen    数字        标题截取长度
'picadd   all/att/http   空白或all,调所有图片/ att，只调上传在“attachments”下的图片/ http，只调转帖的图片
'
'●act取值说明：
'new	        新发主题 
'newreply		新回主题 
'hotr			热贴,即按点击量排行
'hotc(或hot)  	人气贴,即按回复量排行
'top			置顶贴
'recommend		推荐贴或精华贴
'----------------------------------------------------------------------------
%>
<%
Sub dyfPic(f_width,f_height,t_height,act,forumid,number,TitleLen,picadd)
f_width=replace(trim(f_width),"'","")
  if f_width=""  then f_width=300
f_height=replace(trim(f_height),"'","")
  if f_height=""  then f_height=208
t_height=replace(trim(t_height),"'","")
  if t_height=""  then t_height=19
%>
<script language="javascript">
	var focus_width=<%=f_width%>     /*FLASH幻灯图片宽度*/
	var focus_height=<%=f_height%>  /*FLASH幻灯图片高度*/
	var text_height=<%=t_height%>    /*FLASH幻灯文字标题高度*/
	
	var swf_height = focus_height+text_height		
	var pics = '';
	var links = '';
	var texts = '';							
	function ati(url, img, title)
	{
		if(pics != '')
		{
		pics = "|" + pics;
		links = "|" + links;
		texts = "|" + texts;
		}
		pics = escape(img) + pics;
		links = escape(url) + links;
		texts = title + texts;
	}
<%
call flashPic(act,forumid,number,TitleLen,picadd)
%>
document.write('<embed src="include/flashpic.swf"  wmode="opaque" FlashVars="pics='+pics+'&links='+links+'&texts='+texts+'&borderwidth='+focus_width+'&borderheight='+focus_height+'&textheight='+text_height+'" menu="false" bgcolor="#DADADA" quality="high" width="'+ focus_width +'" height="'+ swf_height +'" allowScriptAccess="sameDomain" type="application/x-shockwave-flash"/>');
</script>
<%
End Sub
%>
<%
sub flashPic(act,forumid,number,TitleLen,picadd)
dim wherestr,orderstr,cycurl,pic_dy,cid
forumid=replace(replace(replace(trim(forumid),"'",""),"，",","),"|",",")
number=Cnum(sqlshow(trim(number)))
TitleLen=Cnum(sqlshow(trim(TitleLen)))
act=sqlshow(trim(act))
picadd=lcase(replace(trim(picadd),"'",""))
  if number="" or number<=0 Or  number>6  then number=6
Select Case picadd
  Case "att" 
      picadd="attachments"
  Case "http" 
      picadd="http://"
  Case Else
      picadd=""
End Select
if forumid="" or forumid="all" then 
    wherestr="classID<998"
else
   wherestr="classID in("&forumid&")"
end if
Select Case Act
	Case "recommend" 
		wherestr="Classic=True and  "&wherestr
		orderstr="order by lastupdatetime DESC"
	Case "new"
      orderstr="order by posttime DESC"
	Case "newreply"
		orderstr="order by lastupdatetime DESC"
    Case "hot" 
      orderstr="order by clicknum DESC"
	Case "hotc" 
		orderstr="order by clicknum DESC"
	Case "hotr" 
		orderstr="order by replynum DESC"
	Case "top" 
		orderstr="order by SetTopAll desc, SetTop DESC,lastupdatetime"
    Case else
		orderstr="order by lastupdatetime DESC"
End Select
Sql="Select Top "&number&" * From content WHERE "&wherestr&"  and content like '%[[]img]"&picadd&"%.jpg[[]/img]%'    "&orderstr&",id DESC"
Openrs(sql)
if Rs.eof and Rs.bof then
Response.Write"ati('./', '../images/nopic.jpg', '没有JPG格式图片可调用');"
Else
do while not Rs.eof
pic_dy=lcase(HTMLEncode(Rs("Content")))
pic_dy=left(pic_dy,instr(pic_dy,"[/img]")-1)
pic_dy=right(pic_dy,len(pic_dy)-4-instr(pic_dy,"[img]"))
If Rs("parent")=0 Then cid=Rs("ID") Else cid=Rs("parent")&CHR(38)&"listMethod=search#"&Rs("ID")
Response.Write"ati('contextsingle.asp?id="&cid&"', '"&pic_dy&"', '"&Left(HTMLEncode(Rs("title")),TitleLen)&"');"& vbCrLf
Rs.movenext
loop
end If
Rs.Close
end Sub
%>