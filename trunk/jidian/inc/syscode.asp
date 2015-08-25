<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="ubbcode.asp"-->
<!--#include file="function.asp"-->


<%

'=================================================
'作  用：转换html函数
'参  考：delHtml(rsnews("contont"),14)
'=================================================
Function delHtml(strHtml) '做了一个函数名叫delhtml

Dim objRegExp, strOutput
Set objRegExp = New Regexp ' 建立正则表达式

objRegExp.IgnoreCase = True ' 设置是否区分大小写
objRegExp.Global = True '是匹配所有字符串还是只是第一个
objRegExp.Pattern = "(<[a-zA-Z].*?>)|(<[\/][a-zA-Z].*?>)" ' 设置模式引号中的是正则表达式，用来找出html标签

strOutput = objRegExp.Replace(strHtml, "") '将html标签去掉
strOutput = Replace(strOutput, "<", "<") '防止非html标签不显示
strOutput = Replace(strOutput, ">", ">") 
delHtml = strOutput

Set objRegExp = Nothing
End Function

'srt1是你要去除html代码字符串，可以其它任何地方读取过来。
str1 = "<meta http-equiv=""refresh"" content=""0;URL=apple/default.htm""><title>正</3>在转到 ... ...</title>" 
'应用函数


'=================================================
'作  用：截取字符函数
'参  考：cutstr(rsnews("title"),14)
'=================================================
function cutstr(tempstr,tempwid)
if len(tempstr)>tempwid then
cutstr=left(tempstr,tempwid)&"..."
else
cutstr=tempstr
end if
end function
 
dim strFileName,MaxPerPage,ShowSmallClassType
dim totalPut,CurrentPage,TotalPages,tolPut
dim BeginTime,EndTime
dim founderr, errmsg
dim BigClassName,SmallClassName,SpecialName,keyword,keywd,nbtyped
dim rs,rsa,sql,sqlArticle,rsArticle,sqlSearch,rsSearch,sqlBigClass,rsBigClass,sqlSpecial,rsSpecial,rsSch,sqlSch
dim SpecialTotal
BeginTime=Timer
BigClassName=Trim(request("BigClassName"))
SmallClassName=Trim(request("SmallClassName"))
SpecialName=trim(request("SpecialName"))
nbtyped=trim(request("nbtyped"))
keyword=trim(request("keyword"))
if keyword<>"" then 
	keyword=replace(replace(replace(replace(keyword,"'","‘"),"<","&lt;"),">","&gt;")," ","&nbsp;")
end if

keywd=trim(request("keywd"))
if keywd<>"" then 
	keywd=replace(replace(replace(replace(keywd,"'","‘"),"<","&lt;"),">","&gt;")," ","&nbsp;")
end if

if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if

sqlBigClass="select * from BigClass order by BigClassID"
Set rsBigClass= Server.CreateObject("ADODB.Recordset")
rsBigClass.open sqlBigClass,conn,1,1


'=================================================
'过程名：ShowSmallClass_Tree
'作  用：树形目录方式显示栏目
'参  数：无
'=================================================
sub ShowSmallClass_Tree()
	if rsBigClass.bof and rsBigClass.eof then 
		response.Write "The column is under construction……" 'The column is under construction 栏目正在建设中
	else
		dim sqlClass,rsClass,strTree,BigClassNum,i,j
		rsBigClass.movefirst
		BigClassNum=rsBigClass.recordcount
		i=1
		do while not rsBigClass.eof
			if i<BigClassNum then
				strTree=""
			else
				strTree=""
			end if
			sqlClass="select * from SmallClass where BigClassName='" & rsBigClass("BigClassName") & "' Order by SmallClassID"
			Set rsClass= Server.CreateObject("ADODB.Recordset")
			rsClass.open sqlClass,conn,1,1
				strTree= strTree & "<table width=180 border=0 cellpadding=0 cellspacing=0>"
				strTree= strTree & "<tr>"
				strTree= strTree & "<td width=24 height=22>"
				strTree= strTree & "<div align=center><img src=Img/arrow_2.gif /></div></td>"
				strTree= strTree & "<td width=150>"
				strTree= strTree & "<a href='article.asp?BigClassName=" & rsBigClass("BigClassName") & "'>"
				 	strTree=strTree & rsBigClass("BigClassName")
				strTree=strTree & "</a></td>"
				'strTree=strTree & "</td>"
				strTree= strTree & "</tr><tr>"
				strTree=strTree & "<TD height=1 colspan=2 background=img/naSzarym.gif>"
				strTree=strTree & "<IMG height=1 src=img/1x1_pix.gif width=10></TD>"
				strTree=strTree & "</TR>"
				strTree=strTree & "</table>"
				SmallClassNum=rsClass.recordcount
				j=1
				do while not rsClass.eof
					rsClass.movenext
					j=j+1
				loop
			response.write strTree
			rsBigClass.movenext
			i=i+1
		loop
		rsClass.close
		set rsClass=nothing
	end if
end sub


'=================================================
'过程名：ShowClassGuide
'作  用：显示栏目导航位置
'参  数：无
'=================================================
sub ShowClassGuide()
	response.write  "|&nbsp;<a href='article.asp'>网络课程</a>&nbsp;&gt;&gt;"
	if BigClassName="" and SmallClassName="" then
		response.write "&nbsp;所有课程"
	else
		if BigClassName<>"" then
			response.write "&nbsp;<a href='article.asp?BigClassName=" & Server.URLEncode(BigClassName) & "'>" & BigClassName & "</a>&nbsp;&gt;&gt;&nbsp;"
			if SmallClassName<>"" then
				response.write "<a href='article.asp?BigClassName=" & Server.URLEncode(BigClassName) & "&SmallClassName=" & Server.URLEncode(SmallClassName) & "'>" & SmallClassName & "</a>"
			else
				response.write "所有小类"
			end if
		end if
	end if
end sub

'=================================================
'过程名：ShowArticleTotal
'作  用：显示文章总数
'参  数：无
'=================================================
sub ShowArticleTotal()
	dim sqlTotal
	dim rsTotal
	sqlTotal="select Count(*) from Product where Passed=True "
	if BigClassName<>"" then
		sqlTotal=sqlTotal & " and BigClassName='" & BigClassName & "' "
		if SmallClassName<>"" then
			sqlTotal=sqlTotal & " and SmallClassName='" & SmallClassName & "' "
		end if
	else
		if SpecialName<>"" then
			sqlTotal=sqlTotal & " and SpecialName='" & SpecialName & "' "
		end if
	end if
	Set rsTotal= Server.CreateObject("ADODB.Recordset")
	rsTotal.open sqlTotal,conn,1,1
	if rsTotal.eof and rsTotal.bof then
		totalPut=0
		response.write "共有 0 条记录"
	else
		totalPut=rsTotal(0)
		response.Write "共有 " & totalPut & " 条记录"
	end if
	rsTotal.close
	set rsTotal=nothing
end sub

'=================================================
'过程名：ShowArticle
'=================================================
sub ShowArticle(TitleLen)
	if TitleLen<0 or TitleLen>200 then
		TitleLen=50
	end if
    if currentpage<1 then
	   	currentpage=1
    end if
	if (currentpage-1)*MaxPerPage>totalput then
		if (totalPut mod MaxPerPage)=0 then
	   		currentpage= totalPut \ MaxPerPage
		else
		   	currentpage= totalPut \ MaxPerPage + 1
		end if
   	end if
	if currentPage=1 then
        sqlArticle="select top " & MaxPerPage	
	else
		sqlArticle="select "
	end if

	sqlArticle=sqlArticle & " * from Product where Passed=True "
	
	if BigClassName<>"" then
		sqlArticle=sqlArticle & " and BigClassName='" & BigClassName & "' "
		if SmallClassName<>"" then
			sqlArticle=sqlArticle & " and SmallClassName='" & SmallClassName & "' "
		end if
	else
		if SpecialName<>"" then
			sqlArticle=sqlArticle & " and SpecialName='" & SpecialName & "' "
		end if
	end if
	sqlArticle=sqlArticle & " order by articleid desc"
	Set rsArticle= Server.CreateObject("ADODB.Recordset")
	rsArticle.open sqlArticle,conn,1,1
	if rsArticle.bof and  rsArticle.eof then
		response.Write("<br><li>没有记录</li>")
	else
		if currentPage=1 then
			call ArticleContent(TitleLen)
		else
			if (currentPage-1)*MaxPerPage<totalPut then
         	   	rsArticle.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rsArticle.bookmark
            	call ArticleContent(TitleLen)
        	else
	        	currentPage=1
           		call ArticleContent(TitleLen)
	    	end if
		end if
	end if
	rsArticle.close
	set rsArticle=nothing
end sub

sub ArticleContent(intTitleLen)
   	dim i,strTemp
	dim sqlClassWrd
dim selclass
	selclass=rsArticle("BigClassName")
sqlClassWrd="select * from BigClass where BigClassName='" & selclass & "'"
Set rsClassWrd= Server.CreateObject("ADODB.Recordset")
rsClassWrd.open sqlClassWrd,conn,1,1
	
    'response.Write "<table width='100%' border='1' bordercolor='#BFC9D1' style='border-collapse:collapse;'>"
    'response.Write "<tr id='chplisttd'><td>文章类别文章标题更新时间</td></tr>" 
	'if rsClassWrd("Wrd1")<>"" then response.write "<td>" & rsClassWrd("Wrd1") & "</td>" else response.write" " end if 
	response.Write  "<ul>"
    i=0
	do while not rsArticle.eof
		strTemp=""
		'strTemp = strTemp & ""
			strTemp= strTemp & "<li>"
			strTemp= strTemp & "<a class='arttype' href='article.asp?BigClassName=" & rsArticle("BigClassName") & "&SmallClassName=" & rsArticle("SmallClassName") & "'>【" & rsArticle("SmallClassName") & "】</a> " 
			strTemp= strTemp & "<span class='artlnk'><a href=ArticleShow.asp?ArticleID=" & rsArticle("articleid") & ">" 
			strTemp= strTemp & rsArticle("Title")
			strTemp= strTemp & "</a></span> "
			
			if datediff("d",rsArticle("UpdateTime"),date())<3 then
            strTemp= strTemp & " <span class='newest' title='最新发布'></span> "
            else
            strTemp= strTemp & " " 
            end if
            if rsArticle("Elite")=True then 
            strTemp= strTemp & "<span class='elite' title='推荐课程'></span>"
			else 
			strTemp= strTemp & " " 
		    end if	
			if rsArticle("Hits")>5 then 
            strTemp= strTemp & "<span class='hot' title='最热课程'></span>"
			else 
			strTemp= strTemp & " " 
		    end if

			strTemp= strTemp & "<span class='artlistupdate'>" 
			strTemp= strTemp & rsArticle("UpdateTime")
            strTemp= strTemp & "</span>"
			strTemp= strTemp & "</li>"
			
'           if rsClassWrd("Wrd1")<>"" then 
'           strTemp= strTemp & "<td style='padding: 4px;word-break:break-all;word-wrap : break-word;'>"
'           strTemp= strTemp & rsArticle("Zi1") 
'			strTemp= strTemp & "</td>"
'			else 
'			strTemp= strTemp & " " 
'		    end if	
			
'			strTemp= strTemp & "<td chpnumber=18>编号："
'           strTemp= strTemp & rsArticle("Product_Id")  
'			strTemp= strTemp & "</td>"

'			strTemp= strTemp & "<tr>" 
'			strTemp= strTemp & "<td height=1 colspan=5 bgcolor=#CCCCCC></td>"
'			strTemp= strTemp & "</tr>"
'			strTemp= strTemp & "</table>"
		response.write strTemp
		rsArticle.movenext
		i=i+1
		if i>=MaxPerPage then exit do	
	loop
    'response.Write "</table>"
	response.Write "</ul>"
end sub 
    

'=================================================
'过程名：ShowSearchTerm
'作  用：显示搜索条件信息
'参  数：无
'=================================================
sub ShowSearchTerm()
	response.write "|&nbsp;学习园地文章搜索&nbsp;&gt;&gt; "
	if BigClassName<>"" then
		response.write "<a href='article.asp?BigClassName=" & BigClassName & "'>" & BigClassName & "</a>&nbsp;&gt;&gt;&nbsp;"
		if SmallClassName<>"" then
			response.write "<a href='article.asp?BigClassName=" & BigClassName & "&SmallClassName=" & SmallClassName & "'>" & SmallClassName & "</a>&nbsp;&gt;&gt;&nbsp;"
		else
			response.write "所有小类&nbsp;&gt;&gt;&nbsp;"
		end if
	end if
	if keyword<>"" then
	  response.Write "含有 <font color=red>"&keyword&"</font> 的信息"
	else
	  response.write "&nbsp;所有信息"
	end if
end sub

'=================================================
'过程名：SearchResultTotal
'作  用：显示搜索结果总数
'参  数：无
'=================================================
sub SearchResultTotal()
	dim rsTotal,sqlTotal
	sqlTotal="select count(*) from Product where Passed=True "
	if BigClassName<>"" then
		sqlTotal=sqlTotal & " and BigClassName='" & BigClassName & "' "
		if SmallClassName<>"" then
			sqlTotal=sqlTotal & " and SmallClassName='" & SmallClassName & "' "
		end if
	else
		if SpecialName<>"" then
			sqlTotal=sqlTotal & " and SpecialName='" & SpecialName & "' "
		end if
	end if
	if keyword<>"" then
		sqlTotal=sqlTotal & " and (Title like '%" & keyword & "%' or Content like '%" & keyword & "%' )"
	end if
	Set rsTotal= Server.CreateObject("ADODB.Recordset")
	rsTotal.open sqlTotal,conn,1,1
	if rsTotal.eof and rsTotal.bof then
		totalPut=0
		response.write "共有 0 条信息"
	else
		totalPut=rsTotal(0)
		response.Write "共找到 " & totalPut & " 条信息"
	end if
end sub

'=================================================
'过程名：ShowSearchResult
'作  用：分页显示搜索结果
'参  数：无
'=================================================
sub ShowSearchResult()
    if currentpage<1 then
	   	currentpage=1
    end if
	if (currentpage-1)*MaxPerPage>totalput then
		if (totalPut mod MaxPerPage)=0 then
	   		currentpage= totalPut \ MaxPerPage
		else
		   	currentpage= totalPut \ MaxPerPage + 1
		end if
   	end if
	if currentPage=1 then
        sqlSearch="select top " & MaxPerPage	
	else
		sqlSearch="select "
	end if

	sqlSearch=sqlSearch & " * from Product where Passed=True "
	if BigClassName<>"" then
		sqlSearch=sqlSearch & " and BigClassName='" & BigClassName & "' "
		if SmallClassName<>"" then
			sqlSearch=sqlSearch & " and SmallClassName='" & SmallClassName & "' "
		end if
	else
		if SpecialName<>"" then
			sqlSearch=sqlSearch & " and SpecialName='" & SpecialName & "' "
		end if
	end if
	if keyword<>"" then
		sqlSearch=sqlSearch & " and (Title like '%" & keyword & "%' or Content like '%"& keyword &"%' )"
	end if
	sqlSearch=sqlSearch & " order by articleid desc"
	Set rsSearch= Server.CreateObject("ADODB.Recordset")
	rsSearch.open sqlSearch,conn,1,1
 	if rsSearch.eof and rsSearch.bof then 
       		response.write "<p align='center'><br><br>没有或没有找到相关信息</p>" 
   	else 
   		if currentPage=1 then 
       		call SearchResultContent()
   		else 
       		if (currentPage-1)*MaxPerPage<totalPut then 
       			rsSearch.move  (currentPage-1)*MaxPerPage 
       			dim bookmark 
       			bookmark=rsSearch.bookmark 
       			call SearchResultContent()
      		else 
        		currentPage=1 
       			call SearchResultContent()
      		end if 
	   	end if 
   	end if 
   	rsSearch.close 
   	set rsSearch=nothing   
end sub

sub SearchResultContent()
   	dim i,strTemp,Content
	i=1
	do while not rsSearch.eof
		strTemp=""
		strTemp=strTemp & cstr(i) & ".<a href='ArticleShow.asp?ArticleID=" & rsSearch("articleid") & "'>"
			strTemp=strTemp & "<b>" & replace(rsSearch("title"),""&keyword&"","<font color=red>"&keyword&"</font>") & "</b></font></a>"
		strTemp=strTemp & " [" & FormatDateTime(rsSearch("UpdateTime"),1) & "]"
		Content=left(nohtml(rsSearch("Content")),200)
			strTemp=strTemp & "<div style='padding:10px 20px'>" & replace(Content,""&keyword&"","<font color=red>"&keyword&"</font>") & "……</div>"
		'strTemp=strTemp & "</a>"
		response.write strTemp
		i=i+1
		if i>MaxPerPage then exit do
		rsSearch.movenext
	loop
end sub 


'=================================================
'过程名：ShowSearch
'作  用：显示文章搜索表单
'参  数：ShowType ----显示方式。1为纵向，2为横向
'=================================================
sub ShowSearch(ShowType)
	dim count
	if ShowType<>1 and ShowType<>2 then
		ShowType=1
	end if
	
if ShowType=1 then%>
<div>
<%end if%>
<form method="Get" name="myform" action="search.asp">
<input type="text" name="keyword"  value="请输入学习文章关键字" maxlength="50" onBlur="if(this.value==''){this.value='请输入学习文章关键字'};" 
            onfocus="if(this.value=='请输入学习文章关键字'){this.value=''};"  class="schipt">
<%if ShowType=1 then%>
</div>
<div>
<%end if%>
<input type="submit" name="Submit"  value="搜 索" class="schbtn">
</form>
<%if ShowType=1 then%>
</div>
<%end if%>



<%
end sub

'=================================================
'过程名：ShowAllClass
'作  用：显示所有栏目（栏目导航）
'参  数：无
'=================================================
sub ShowAllClass()
	if rsBigClass.bof and rsBigClass.eof then 
		response.Write "&nbsp;没有任何栏目"
	else
		dim sqlClass,rsClass,strClassName
		rsBigClass.movefirst
		do while not rsBigClass.eof
			strClassName= "<div class='rbigclass'><a href='article.asp?BigClassName=" & Server.URLEncode(rsBigClass("BigClassName")) & "'>" & rsBigClass("BigClassName") & "</a></div><div class='rBrief'>" & rsBigClass("Brief") & "</div>"
'			sqlClass="select * from SmallClass where BigClassName='" & rsBigClass("BigClassName") & "' Order by SmallClassID"
'			Set rsClass= Server.CreateObject("ADODB.Recordset")
'			rsClass.open sqlClass,conn,1,1
'			do while not rsClass.eof
'				strClassName=strClassName & "<span class='rsmallclass'><a href='article.asp?BigClassName=" & Server.URLEncode(rsClass("BigClassName")) & "&SmallClassName=" & Server.URLEncode(rsClass("SmallClassName")) & "'>" & rsClass("SmallClassName") & "</a></span> | "
'				rsClass.movenext
'			loop
		response.write strClassName & "<p></p>"
		rsBigClass.movenext
		loop
		rsClass.close
		set rsClass=nothing
	end if
end sub


'=================================================
'过程名：ShowMenu
'作  用：显示大类栏目（下拉菜单）
'参  数：无
'=================================================
sub ShowMenu()
	if rsBigClass.bof and rsBigClass.eof then 
		response.Write "没有任何栏目"
	else
		dim rsClass,strClassName
		rsBigClass.movefirst
		response.write strClassName & "<ul>"
		do while not rsBigClass.eof
			response.write strClassName & "<li><a href='article.asp?BigClassName=" & rsBigClass("BigClassName") & "'>" & rsBigClass("BigClassName") & "</a>"
			response.write strClassName & "</li>"
			rsBigClass.movenext
		loop
		response.write strClassName & "</ul>"
		rsClass.close
		set rsClass=nothing
	end if
end sub


'=================================================
'过程名：ShowAllClass1
'作  用：显示所有文章类别（文章导航树型收缩菜单）
'参  数：无
'=================================================
sub ShowAllClassl()
%>
  <%	dim m,n
						m=1
						set rsl=server.createobject("adodb.recordset")
					 	sql="select * from BigClass order by BigClassID "
						rsl.open sql,conn,1,1
						if rsl.eof and rsl.bof then
						response.write""
						else
						do while not rsl.eof
						m=m+1
						n=rsl("BigClassName")%>
  <% set rs2=server.createobject("adodb.recordset")
							sql2="select * from SmallClass where BigClassName='"&n&"' order by SmallClassID"
							rs2.open sql2,conn,1,1
							if rs2.eof or rs2.bof then	%>
  <div class=b_title style="cursor:hand;" onmouseover=this.className='b_title2'; onmouseout=this.className='b_title';> <a href="article.asp?BigClassName=<%=Server.URLEncode(n)%>"><%=n%></a></div>
  <%else%>
  <%if  BigClassName=rs2("BigClassName") then%>   
  <div class=b_title2> <%=rs2("BigClassName")%></div>
        <%
		  do while not rs2.eof
		%>
        <div class=s_title style="cursor:hand;" onmouseover=this.className='s_title2'; onmouseout=this.className='s_title';><a href="article.asp?BigClassName=<%=Server.URLEncode(n)%>&SmallClassName=<%=Server.URLEncode(rs2("SmallClassName"))%>"><%=rs2("SmallClassName")%></a></div>
        <%rs2.movenext
							loop							
							rs2.close
							set rs2=nothing%>
  <%else%>
    <div class=b_title style="cursor:hand;" onclick="showsubmenu(<%=m%>);window.location='article.asp?BigClassName=<%=Server.URLEncode(n)%>'" onmouseover=this.className='b_title2'; onmouseout=this.className='b_title';> <%=rs2("BigClassName")%></div>
    <div id='submenu<%=m%>' style="display:none"> 
        <%
			do while not rs2.eof
							%>
        <div class=s_title style="cursor:hand;" onmouseover=this.className='s_title2'; onmouseout=this.className='s_title';><a href="article.asp?BigClassName=<%=Server.URLEncode(n)%>&SmallClassName=<%=Server.URLEncode(rs2("SmallClassName"))%>"><%=rs2("SmallClassName")%></a></div>
        <%rs2.movenext
							loop							
							rs2.close
							set rs2=nothing%>
  </div> 
  <%end if%> <%end if%> 
   <%rsl.movenext
					loop
					rsl.close
					set rsl=nothing
					end if%>
					

<%
end sub

'=================================================
'过程名：ShowBigclass
'作  用：显示大类栏目（下拉菜单）
'参  数：无
'=================================================
sub ShowBigClass()
	if rsBigClass.bof and rsBigClass.eof then 
		response.Write "没有任何栏目"
	else
		dim rsClass,strClassName
		rsBigClass.movefirst
		response.write strClassName & "<ul>"
		do while not rsBigClass.eof
			response.write strClassName & "<li id='one'><a href='article.asp?BigClassName=" & rsBigClass("BigClassName") & "'>" & rsBigClass("BigClassName") & "</a>"
			response.write strClassName & "</li>"
			rsBigClass.movenext
		loop
		response.write strClassName & "</ul>"
		rsClass.close
		set rsClass=nothing
	end if
end sub



'=================================================
'过程名：ShowArticleContent
'作  用：显示文章具体的内容，可以分页显示
'参  数：无
'"<img border=0 src=" & rsArticle("DefaultPicUrl") & " width=100 height=120>"
'=================================================
sub ShowArticleContent()
	dim ArticleID,strContent,CurrentPage
	dim ContentLen,MaxPerPage,pages,i,lngBound
	dim BeginPoint,EndPoint
	ArticleID=rs("ArticleID")
	strContent=rs("Content")
	ContentLen=len(strContent)
	CurrentPage=trim(request("ArticlePage"))
	if ShowContentByPage="No" or ContentLen<=200000 then
		response.write strContent
		if ShowContentByPage="Yes" then
			response.write "</p><p align='center'></p>"
		end if
	else
		if CurrentPage="" then
			CurrentPage=1
		else
			CurrentPage=Cint(CurrentPage)
		end if
		pages=ContentLen\MaxPerPage_Content
		if MaxPerPage_Content*pages<ContentLen then
			pages=pages+1
		end if
		lngBound=MaxPerPage_Content          '最大误差范围
		if CurrentPage<1 then CurrentPage=1
		if CurrentPage>pages then CurrentPage=pages

		dim lngTemp
		dim lngTemp1,lngTemp1_1,lngTemp1_2,lngTemp1_1_1,lngTemp1_1_2,lngTemp1_1_3,lngTemp1_2_1,lngTemp1_2_2,lngTemp1_2_3
		dim lngTemp2,lngTemp2_1,lngTemp2_2,lngTemp2_1_1,lngTemp2_1_2,lngTemp2_2_1,lngTemp2_2_2
		dim lngTemp3,lngTemp3_1,lngTemp3_2,lngTemp3_1_1,lngTemp3_1_2,lngTemp3_2_1,lngTemp3_2_2
		dim lngTemp4,lngTemp4_1,lngTemp4_2,lngTemp4_1_1,lngTemp4_1_2,lngTemp4_2_1,lngTemp4_2_2
		dim lngTemp5,lngTemp5_1,lngTemp5_2
		dim lngTemp6,lngTemp6_1,lngTemp6_2
		
		if CurrentPage=1 then
			BeginPoint=1
		else
			BeginPoint=MaxPerPage_Content*(CurrentPage-1)+1
			
			lngTemp1_1_1=instr(BeginPoint,strContent,"</table>",1)
			lngTemp1_1_2=instr(BeginPoint,strContent,"</TABLE>",1)
			lngTemp1_1_3=instr(BeginPoint,strContent,"</Table>",1)
			if lngTemp1_1_1>0 then
				lngTemp1_1=lngTemp1_1_1
			elseif lngTemp1_1_2>0 then
				lngTemp1_1=lngTemp1_1_2
			elseif lngTemp1_1_3>0 then
				lngTemp1_1=lngTemp1_1_3
			else
				lngTemp1_1=0
			end if
							
			lngTemp1_2_1=instr(BeginPoint,strContent,"<table",1)
			lngTemp1_2_2=instr(BeginPoint,strContent,"<TABLE",1)
			lngTemp1_2_3=instr(BeginPoint,strContent,"<Table",1)
			if lngTemp1_2_1>0 then
				lngTemp1_2=lngTemp1_2_1
			elseif lngTemp1_2_2>0 then
				lngTemp1_2=lngTemp1_2_2
			elseif lngTemp1_2_3>0 then
				lngTemp1_2=lngTemp1_2_3
			else
				lngTemp1_2=0
			end if
			
			if lngTemp1_1=0 and lngTemp1_2=0 then
				lngTemp1=BeginPoint
			else
				if lngTemp1_1>lngTemp1_2 then
					lngtemp1=lngTemp1_2
				else
					lngTemp1=lngTemp1_1+8
				end if
			end if

			lngTemp2_1_1=instr(BeginPoint,strContent,"</p>",1)
			lngTemp2_1_2=instr(BeginPoint,strContent,"</P>",1)
			if lngTemp2_1_1>0 then
				lngTemp2_1=lngTemp2_1_1
			elseif lngTemp2_1_2>0 then
				lngTemp2_1=lngTemp2_1_2
			else
				lngTemp2_1=0
			end if
						
			lngTemp2_2_1=instr(BeginPoint,strContent,"<p",1)
			lngTemp2_2_2=instr(BeginPoint,strContent,"<P",1)
			if lngTemp2_2_1>0 then
				lngTemp2_2=lngTemp2_2_1
			elseif lngTemp2_2_2>0 then
				lngTemp2_2=lngTemp2_2_2
			else
				lngTemp2_2=0
			end if
			
			if lngTemp2_1=0 and lngTemp2_2=0 then
				lntTemp2=BeginPoint
			else
				if lngTemp2_1>lngTemp2_2 then
					lngtemp2=lngTemp2_2
				else
					lngTemp2=lngTemp2_1+4
				end if
			end if

			lngTemp3_1_1=instr(BeginPoint,strContent,"</ur>",1)
			lngTemp3_1_2=instr(BeginPoint,strContent,"</UR>",1)
			if lngTemp3_1_1>0 then
				lngTemp3_1=lngTemp3_1_1
			elseif lngTemp3_1_2>0 then
				lngTemp3_1=lngTemp3_1_2
			else
				lngTemp3_1=0
			end if
			
			lngTemp3_2_1=instr(BeginPoint,strContent,"<ur",1)
			lngTemp3_2_2=instr(BeginPoint,strContent,"<UR",1)
			if lngTemp3_2_1>0 then
				lngTemp3_2=lngTemp3_2_1
			elseif lngTemp3_2_2>0 then
				lngTemp3_2=lngTemp3_2_2
			else
				lngTemp3_2=0
			end if
					
			if lngTemp3_1=0 and lngTemp3_2=0 then
				lngTemp3=BeginPoint
			else
				if lngTemp3_1>lngTemp3_2 then
					lngtemp3=lngTemp3_2
				else
					lngTemp3=lngTemp3_1+5
				end if
			end if
			
			if lngTemp1<lngTemp2 then
				lngTemp=lngTemp2
			else
				lngTemp=lngTemp1
			end if
			if lngTemp<lngTemp3 then
				lngTemp=lngTemp3
			end if

			if lngTemp>BeginPoint and lngTemp<=BeginPoint+lngBound then
				BeginPoint=lngTemp
			else
				lngTemp4_1_1=instr(BeginPoint,strContent,"</li>",1)
				lngTemp4_1_2=instr(BeginPoint,strContent,"</LI>",1)
				if lngTemp4_1_1>0 then
					lngTemp4_1=lngTemp4_1_1
				elseif lngTemp4_1_2>0 then
					lngTemp4_1=lngTemp4_1_2
				else
					lngTemp4_1=0
				end if
				
				lngTemp4_2_1=instr(BeginPoint,strContent,"<li",1)
				lngTemp4_2_1=instr(BeginPoint,strContent,"<LI",1)
				if lngTemp4_2_1>0 then
					lngTemp4_2=lngTemp4_2_1
				elseif lngTemp4_2_2>0 then
					lngTemp4_2=lngTemp4_2_2
				else
					lngTemp4_2=0
				end if
				
				if lngTemp4_1=0 and lngTemp4_2=0 then
					lngTemp4=BeginPoint
				else
					if lngTemp4_1>lngTemp4_2 then
						lngtemp4=lngTemp4_2
					else
						lngTemp4=lngTemp4_1+5
					end if
				end if
				
				if lngTemp4>BeginPoint and lngTemp4<=BeginPoint+lngBound then
					BeginPoint=lngTemp4
				else					
					lngTemp5_1=instr(BeginPoint,strContent,"<img",1)
					lngTemp5_2=instr(BeginPoint,strContent,"<IMG",1)
					if lngTemp5_1>0 then
						lngTemp5=lngTemp5_1
					elseif lngTemp5_2>0 then
						lngTemp5=lngTemp5_2
					else
						lngTemp5=BeginPoint
					end if
					
					if lngTemp5>BeginPoint and lngTemp5<BeginPoint+lngBound then
						BeginPoint=lngTemp5
					else
						lngTemp6_1=instr(BeginPoint,strContent,"<br>",1)
						lngTemp6_2=instr(BeginPoint,strContent,"<BR>",1)
						if lngTemp6_1>0 then
							lngTemp6=lngTemp6_1
						elseif lngTemp6_2>0 then
							lngTemp6=lngTemp6_2
						else
							lngTemp6=0
						end if
					
						if lngTemp6>BeginPoint and lngTemp6<BeginPoint+lngBound then
							BeginPoint=lngTemp6+4
						end if
					end if
				end if
			end if
		end if

		if CurrentPage=pages then
			EndPoint=ContentLen
		else
		  EndPoint=MaxPerPage_Content*CurrentPage
		  if EndPoint>=ContentLen then
			EndPoint=ContentLen
		  else
			lngTemp1_1_1=instr(EndPoint,strContent,"</table>",1)
			lngTemp1_1_2=instr(EndPoint,strContent,"</TABLE>",1)
			lngTemp1_1_3=instr(EndPoint,strContent,"</Table>",1)
			if lngTemp1_1_1>0 then
				lngTemp1_1=lngTemp1_1_1
			elseif lngTemp1_1_2>0 then
				lngTemp1_1=lngTemp1_1_2
			elseif lngTemp1_1_3>0 then
				lngTemp1_1=lngTemp1_1_3
			else
				lngTemp1_1=0
			end if
							
			lngTemp1_2_1=instr(EndPoint,strContent,"<table",1)
			lngTemp1_2_2=instr(EndPoint,strContent,"<TABLE",1)
			lngTemp1_2_3=instr(EndPoint,strContent,"<Table",1)
			if lngTemp1_2_1>0 then
				lngTemp1_2=lngTemp1_2_1
			elseif lngTemp1_2_2>0 then
				lngTemp1_2=lngTemp1_2_2
			elseif lngTemp1_2_3>0 then
				lngTemp1_2=lngTemp1_2_3
			else
				lngTemp1_2=0
			end if
			
			if lngTemp1_1=0 and lngTemp1_2=0 then
				lngTemp1=EndPoint
			else
				if lngTemp1_1>lngTemp1_2 then
					lngtemp1=lngTemp1_2-1
				else
					lngTemp1=lngTemp1_1+7
				end if
			end if

			lngTemp2_1_1=instr(EndPoint,strContent,"</p>",1)
			lngTemp2_1_2=instr(EndPoint,strContent,"</P>",1)
			if lngTemp2_1_1>0 then
				lngTemp2_1=lngTemp2_1_1
			elseif lngTemp2_1_2>0 then
				lngTemp2_1=lngTemp2_1_2
			else
				lngTemp2_1=0
			end if
						
			lngTemp2_2_1=instr(EndPoint,strContent,"<p",1)
			lngTemp2_2_2=instr(EndPoint,strContent,"<P",1)
			if lngTemp2_2_1>0 then
				lngTemp2_2=lngTemp2_2_1
			elseif lngTemp2_2_2>0 then
				lngTemp2_2=lngTemp2_2_2
			else
				lngTemp2_2=0
			end if
			
			if lngTemp2_1=0 and lngTemp2_2=0 then
				lngTemp2=EndPoint
			else
				if lngTemp2_1>lngTemp2_2 then
					lngTemp2=lngTemp2_2-1
				else
					lngTemp2=lngTemp2_1+3
				end if
			end if

			lngTemp3_1_1=instr(EndPoint,strContent,"</ur>",1)
			lngTemp3_1_2=instr(EndPoint,strContent,"</UR>",1)
			if lngTemp3_1_1>0 then
				lngTemp3_1=lngTemp3_1_1
			elseif lngTemp3_1_2>0 then
				lngTemp3_1=lngTemp3_1_2
			else
				lngTemp3_1=0
			end if
			
			lngTemp3_2_1=instr(EndPoint,strContent,"<ur",1)
			lngTemp3_2_2=instr(EndPoint,strContent,"<UR",1)
			if lngTemp3_2_1>0 then
				lngTemp3_2=lngTemp3_2_1
			elseif lngTemp3_2_2>0 then
				lngTemp3_2=lngTemp3_2_2
			else
				lngTemp3_2=0
			end if
					
			if lngTemp3_1=0 and lngTemp3_2=0 then
				lngTemp3=EndPoint
			else
				if lngTemp3_1>lngTemp3_2 then
					lngtemp3=lngTemp3_2-1
				else
					lngTemp3=lngTemp3_1+4
				end if
			end if
			
			if lngTemp1<lngTemp2 then
				lngTemp=lngTemp2
			else
				lngTemp=lngTemp1
			end if
			if lngTemp<lngTemp3 then
				lngTemp=lngTemp3
			end if

			if lngTemp>EndPoint and lngTemp<=EndPoint+lngBound then
				EndPoint=lngTemp
			else
				lngTemp4_1_1=instr(EndPoint,strContent,"</li>",1)
				lngTemp4_1_2=instr(EndPoint,strContent,"</LI>",1)
				if lngTemp4_1_1>0 then
					lngTemp4_1=lngTemp4_1_1
				elseif lngTemp4_1_2>0 then
					lngTemp4_1=lngTemp4_1_2
				else
					lngTemp4_1=0
				end if
				
				lngTemp4_2_1=instr(EndPoint,strContent,"<li",1)
				lngTemp4_2_1=instr(EndPoint,strContent,"<LI",1)
				if lngTemp4_2_1>0 then
					lngTemp4_2=lngTemp4_2_1
				elseif lngTemp4_2_2>0 then
					lngTemp4_2=lngTemp4_2_2
				else
					lngTemp4_2=0
				end if
				
				if lngTemp4_1=0 and lngTemp4_2=0 then
					lngTemp4=EndPoint
				else
					if lngTemp4_1>lngTemp4_2 then
						lngtemp4=lngTemp4_2-1
					else
						lngTemp4=lngTemp4_1+4
					end if
				end if
				
				if lngTemp4>EndPoint and lngTemp4<=EndPoint+lngBound then
					EndPoint=lngTemp4
				else					
					lngTemp5_1=instr(EndPoint,strContent,"<img",1)
					lngTemp5_2=instr(EndPoint,strContent,"<IMG",1)
					if lngTemp5_1>0 then
						lngTemp5=lngTemp5_1-1
					elseif lngTemp5_2>0 then
						lngTemp5=lngTemp5_2-1
					else
						lngTemp5=EndPoint
					end if
					
					if lngTemp5>EndPoint and lngTemp5<EndPoint+lngBound then
						EndPoint=lngTemp5
					else
						lngTemp6_1=instr(EndPoint,strContent,"<br>",1)
						lngTemp6_2=instr(EndPoint,strContent,"<BR>",1)
						if lngTemp6_1>0 then
							lngTemp6=lngTemp6_1+3
						elseif lngTemp6_2>0 then
							lngTemp6=lngTemp6_2+3
						else
							lngTemp6=EndPoint
						end if
					
						if lngTemp6>EndPoint and lngTemp6<EndPoint+lngBound then
							EndPoint=lngTemp6
						end if
					end if
				end if
			end if
		  end if
		end if
		response.write mid(strContent,BeginPoint,EndPoint-BeginPoint)
		
		response.write "</p><p align='center'><b>"
		if CurrentPage>1 then
			response.write "<a href='ArticleShow.asp?ArticleID=" & ArticleID & "&ArticlePage=" & CurrentPage-1 & "'>上一页</a>&nbsp;&nbsp;"
		end if
		for i=1 to pages
			if i=CurrentPage then
				response.write "<font color='red'>[" & cstr(i) & "]</font>&nbsp;"
			else
				response.write "<a href='ArticleShow.asp?ArticleID=" & ArticleID & "&ArticlePage=" & i & "'>[" & i & "]</a>&nbsp;"
			end if
		next
		if CurrentPage<pages then
			response.write "&nbsp;<a href='ArticleShow.asp?ArticleID=" & ArticleID & "&ArticlePage=" & CurrentPage+1 & "'>下一页</a>"
		end if
		response.write "</b></p>"
	end if

end sub

set rspro=server.createobject("adodb.recordset")
sqlpro="select top 8 *  from ProSet  order by mdy_time desc"
rspro.open sqlpro,conn,1,1
'=================================================
'过程名：ShowProSet
'作  用：显示专业设置（左侧菜单）
'参  数：无
'=================================================
sub ShowProSet()
  If rspro.eof and rspro.bof then
    response.write "暂无内容"
  else 
  dim strnewsName
  rspro.movefirst
  response.write strnewsName & "<ul>"
  do while not rspro.eof
response.write strnewsName & "<li id='one'><a href='proSetInfo.asp?id=" & rspro("id") & "'>" & cutstr(rspro("title"),14) & "</a></li>"
     rspro.movenext
  loop
  response.write strnewsName & "</ul>"
  rspro.close
  set rspro=nothing
end if
end sub

set rscnews=server.createobject("adodb.recordset")
sqlcnews="select top 6 *  from FacResource  order by time desc"
rscnews.open sqlcnews,conn,1,1
'=================================================
'过程名：ShowFacResource
'作  用：显示设施资源（左侧菜单）
'参  数：无
'=================================================
sub ShowFacResource()
  If rscnews.eof and rscnews.bof then
    response.write "<ul><li>暂无项目</li></ul>"
  else 
  dim strnewsName
  rscnews.movefirst
  response.write strnewsName & "<ul>"
  do while not rscnews.eof
     response.write strnewsName & "<li id='one'><a href='facResourceInfo.asp?id=" & rscnews("id") & "'>" & cutstr(rscnews("title"),14) & "</a></li>"
     rscnews.movenext
  loop
  response.write strnewsName & "</ul>"
  rscnews.close
  set rscnews=nothing
end if 
end sub


set rsrnews=server.createobject("adodb.recordset")
sqlrnews="select top 6 *  from Recruit  order by time desc"
rsrnews.open sqlrnews,conn,1,1
'=================================================
'过程名：ShowRecruit
'作  用：显示最新招生就业（左侧菜单）
'参  数：无
'=================================================
sub ShowRecruit()
  If rsrnews.eof and rsrnews.bof then
    response.write "<ul><li>暂无项目</li></ul>"
  else 
  dim strnewsName
  rsrnews.movefirst
  response.write strnewsName & "<ul>"
  do while not rsrnews.eof
     response.write strnewsName & "<li id='one'><a href='recruitInfo.asp?id=" & rsrnews("id") & "'>" & cutstr(rsrnews("title"),14) & "</a></li>"
     rsrnews.movenext
  loop
  response.write strnewsName & "</ul>"
  rsrnews.close
  set rsrnews=nothing
end if 
end sub

%>



<%
'=================================================
'过程名：ShowSchTerm
'作  用：显示搜索条件信息
'参  数：无
'=================================================
sub ShowSchTerm()
	response.write "|&nbsp;文章搜索&nbsp;&gt;&gt; "
	if nbtyped<>"" then
		response.write "<a href='stuZoneInfo.asp?nbtyped=" & nbtyped & "'>" & nbtyped & "</a>&nbsp;&gt;&gt;&nbsp;"
	end if
	if keywd<>"" then
	  response.Write "含有 <font color=red>"&keywd&"</font> 的信息"
	else
	  response.write "&nbsp;所有信息"
	end if
end sub

'=================================================
'过程名：SchResultTotal
'作  用：显示搜索结果总数
'参  数：无
'=================================================
sub SchResultTotal()
	dim rsTol,sqlTol
	sqlTol="select count(*) from StuZone"
	if nbtyped<>"" then
		sqlTol=sqlTol & " where nbtyped='" & nbtyped & "' "
	end if
	if keywd<>"" then
		sqlTol=sqlTol & " where title like '%" & keywd & "%' or content like '%" & keywd & "%' "
	end if
	Set rsTol= Server.CreateObject("ADODB.Recordset")
	rsTol.open sqlTol,conn,1,1
	if rsTol.eof and rsTol.bof then
		tolPut=0
		response.write "共有 0 条信息"
	else
		tolPut=rsTol(0)
		response.Write "共找到 " & tolPut & " 条信息"
	end if
end sub

'=================================================
'过程名：ShowSchResult
'作  用：分页显示搜索结果
'参  数：无
'=================================================
sub ShowSchResult()
    if currentpage<1 then
	   	currentpage=1
    end if
	if (currentpage-1)*MaxPerPage>tolput then
		if (tolPut mod MaxPerPage)=0 then
	   		currentpage= tolPut \ MaxPerPage
		else
		   	currentpage= tolPut \ MaxPerPage + 1
		end if
   	end if
	if currentPage=1 then
        sqlSch="select top " & MaxPerPage	
	else
		sqlSch="select "
	end if

	sqlSch=sqlSch & " * from StuZone"
	if nbtyped<>"" then
		sqlSch=sqlSch & " where nbtyped='" & nbtyped & "' "
	end if
	if keywd<>"" then
		sqlSch=sqlSch & " where title like '%" & keywd & "%' or content like '%"& keywd &"%'"
	end if
	sqlSch=sqlSch & " order by id desc"
	Set rsSch= Server.CreateObject("ADODB.Recordset")
	rsSch.open sqlSch,conn,1,1
 	if rsSch.eof and rsSch.bof then 
       		response.write "<p align='center'><br><br>没有或没有找到任何信息</p>" 
   	else 
   		if currentPage=1 then 
       		call SchResultContent()
   		else 
       		if (currentPage-1)*MaxPerPage<tolPut then 
       			rsSch.move  (currentPage-1)*MaxPerPage 
       			dim bookmark 
       			bookmark=rsSch.bookmark 
       			call SchResultContent()
      		else 
        		currentPage=1 
       			call SchResultContent()
      		end if 
	   	end if 
   	end if 
   	rsSch.close 
   	set rsSch=nothing   
end sub

sub SchResultContent()
   	dim i,strTemp,content
	i=1
	do while not rsSch.eof
		strTemp=""
		strTemp=strTemp & cstr(i) & ".<a href='stuZoneInfo.asp?id=" & rsSch("id") & "'>"
			strTemp=strTemp & "<b>" & replace(rsSch("title"),""&keywd&"","<font color=red>"&keywd&"</font>") & "</b></font></a>"
		strTemp=strTemp & " [" & FormatDateTime(rsSch("UpdateTime"),1) & "]"
		content=left(nohtml(rsSch("content")),200)
			strTemp=strTemp & "<div style='padding:10px 20px'>" & replace(content,""&keywd&"","<font color=red>"&keywd&"</font>") & "……</div>"
		'strTemp=strTemp & "</a>"
		response.write strTemp
		i=i+1
		if i>MaxPerPage then exit do
		rsSch.movenext
	loop
end sub 


'=================================================
'过程名：ShowSch
'作  用：显示文章搜索表单
'参  数：ShowType ----显示方式。1为纵向，2为横向
'=================================================
sub ShowSch(ShowTyp)
	dim count
%>

<form method="Get" name="myfrm" action="sch.asp">
<input type="text" name="keywd"  value="请输入关键字" maxlength="50" onBlur="if(this.value==''){this.value='请输入关键字'};" 
            onfocus="if(this.value=='请输入关键字'){this.value=''};"  class="schipt">

<input type="submit" name="Submit"  value="搜 索" class="schbtn">
</form>
<%end sub%>
