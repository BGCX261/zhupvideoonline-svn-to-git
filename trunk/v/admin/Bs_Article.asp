<!--#include file="bsconfig.asp"-->
<!--#include file="../Inc/Ubbcode.asp"-->
<!--#include file="Inc/Function.asp"-->

<%
dim strFileName
const MaxPerPage=20
dim totalPut,CurrentPage,TotalPages
dim i,j
dim ArticleID
dim Title
dim sql,rs
dim BigClassName,SmallClassName,SpecialName

strFileName="Bs_Article.asp?BigClassName=" & BigClassName & "&SmallClassName=" & SmallClassName & "&SpecialName=" & SpecialName

Title=Trim(request("Title"))
ArticleID=Request("ArticleID")
BigClassName=Trim(request("BigClassName"))
SmallClassName=Trim(request("SmallClassName"))
SpecialName=trim(request("SpecialName"))

if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if

sql="select * from Product where ArticleID>0"
if session("adminRole")<>"test" then
	sql=sql & ""
else
    sql=sql & " and Author='" & Session("adminName") & "' "
end if
if Title<>"" then
	sql=sql & " and title like '%" & Title & "%' "
end if
if BigClassName<>"" then
	sql=sql & " and BigClassName='" & BigClassName & "' "
	if SmallClassName<>"" then
		sql=sql & " and SmallClassName='" & SmallClassName & "' "
	end if
else
	if SpecialName<>"" then
		sql=sql & " and SpecialName='" & SpecialName & "' "
	end if
end if
sql=sql & " order by articleid desc"
Set rs= Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,1
%>
<SCRIPT language=javascript>
function unselectall()
{
    if(document.del.chkAll.checked){
	document.del.chkAll.checked = document.del.chkAll.checked&0;
    } 	
}

function CheckAll(form)
  {
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.Name != "chkAll")
       e.checked = form.chkAll.checked;
    }
  }
function ConfirmDel()
{
   if(confirm("确定要删除选中的产品吗？一旦删除将不能恢复！"))
     return true;
   else
     return false;
	 
}

</SCRIPT>
<!-- #include file="Inc/Head.asp" -->
<%call checkmanage("17")%>
<table align="center" class="a2">
	<tr>
		<td class="a1">视 频 文 件 上 传 管 理</td>
	</tr>
	<tr class="a4">
		<td>
      <table width="100%">
        <tr> 
          <td bgcolor="#C0C0C0" height="25">|&nbsp; 
            <%
dim sqlBigClass,sqlSmallClass,rsBigClass,rsSmallClass,sqlSpecial,rsSpecial
sqlBigClass="select * from BigClass"
Set rsBigClass= Server.CreateObject("ADODB.Recordset")
rsBigClass.open sqlBigClass,conn,1,1
if rsBigClass.eof then 
	response.Write("还没有任何分类，请首先添加分类。")
end if
do while not rsBigClass.eof
	if Server.URLEncode(rsBigClass("BigClassName"))=BigClassName then
		response.Write("<a href='Bs_Article.asp?BigClassName=" & Server.URLEncode(rsBigClass("BigClassName")) & "'><font color='red'>" & rsBigClass("BigClassName") & "</font></a> | ")
	else
		response.Write("<a href='Bs_Article.asp?BigClassName=" & Server.URLEncode(rsBigClass("BigClassName")) & "'>" & rsBigClass("BigClassName") & "</a> | ")
	end if
	rsBigClass.movenext
loop
rsBigClass.close
set rsBigClass=nothing
%>
          </td>
        </tr>
        <%
if BigClassName<>"" then
	sqlSmallClass="select * from SmallClass where BigClassName='" & BigClassName & "'"
	Set rsSmallClass= Server.CreateObject("ADODB.Recordset")
	rsSmallClass.open sqlSmallClass,conn,1,1
	if not (rsSmallClass.bof and rsSmallClass.eof) then
		response.write "<tr class='tdbg'><td bgcolor='#C0C0C0'>"
		do while not rsSmallClass.eof
			if rsSmallClass("SmallClassName")=SmallClassName then
				response.Write("&nbsp;<a href='Bs_Article.asp?BigClassName=" & Server.URLEncode(rsSmallClass("BigClassName")) & "&SmallClassName=" & Server.URLEncode(rsSmallClass("SmallClassName")) & "'><font color='red'>" & rsSmallClass("SmallClassName") & "</font></a>&nbsp;&nbsp;")
			else
				response.Write("&nbsp;<a href='Bs_Article.asp?BigClassName=" & Server.URLEncode(rsSmallClass("BigClassName")) & "&SmallClassName=" & Server.URLEncode(rsSmallClass("SmallClassName")) & "'>" & rsSmallClass("SmallClassName") & "</a>&nbsp;&nbsp;")
			end if
			rsSmallClass.movenext
		loop
		response.write "</td></tr>"
	end if
	rsSmallClass.close
	set rsSmallClass=nothing
end if
%>
      </table>
      <form name="del" method="Post" action="Bs_Article_Del.asp" onSubmit="return ConfirmDel();">
        <table width="100%">
          <tr> 
            <td height="25" bgcolor="#CCCCCC"><a href="Bs_Article.asp">视频作品管理</a> 
              &gt;&gt; 
              <%
if request.querystring="" then
	response.write "所有产品"
else
	if request("Query")<>"" then
		if Title<>"" then
			response.write "名称中含有“<font color=blue>" & Title & "</font>”的视频"
		else
			response.Write("所有视频")
		end if
 	else
		if BigClassName<>"" then
			response.write "<a href='Bs_Article.asp?BigClassName=" & Server.URLEncode(BigClassName) & "'>" & BigClassName & "</a>&nbsp;&gt;&gt;&nbsp;"
			if SmallClassName<>"" then
				response.write "<a href='Bs_Article.asp?BigClassName=" & Server.URLEncode(BigClassName) & "&SmallClassName=" & Server.URLEncode(SmallClassName) & "'>" & SmallClassName & "</a>"
			else
				response.write "所有小类"
			end if
		end if
		if SpecialName<>"" then
			response.write "<font color=red>[专题]</font> " & SpecialName
		end if
	end if
end if
%>
            </td>
            <td width="150" bgcolor="#CCCCCC">&nbsp;
            <%
  	if rs.eof and rs.bof then
		response.write "共找到 0 个视频</td></tr></table>"
	else
    	totalPut=rs.recordcount
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
		response.Write "共找到 " & totalPut & " 个视频"
%>
            </td>
          </tr>
        </table>
        <%		
	    if currentPage=1 then
        	showContent
        	showpage strFileName,totalput,MaxPerPage,true,false,"个视频"
   	 	else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rs.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rs.bookmark
            	showContent
            	showpage strFileName,totalput,MaxPerPage,true,false,"个视频"
        	else
	        	currentPage=1
           		showContent
           		showpage strFileName,totalput,MaxPerPage,true,false,"个视频"
	    	end if
		end if
	end if
%>
        <%  
sub showContent
   	dim i
    i=0
%>
        <table width="100%"  style="word-break:break-all">
          <tr class="opttitle"> 
            <td>选中</td>
            <td>ID</td>
            <td>编号</td>
            <td>上传者</td>
            <td>作品名称</td>
            <td>加入时间</td>
            <td>访问/回复</td>
            <td>审核情况</td>
            <td>操作</td>
          </tr>
          <%do while not rs.eof%>
          <tr class="opt"> 
            <td> <input name='ArticleID' type='checkbox' onClick="unselectall()" id="ArticleID" value='<%=cstr(rs("articleID"))%>'> 
            </td>
            <td><%=rs("articleid")%></td>
            <td><%=rs("Product_Id")%></td>
            <td><%=rs("Author")%></td>
            <td><a href="../ArticleShow.asp?ArticleID=<%=rs("articleid")%>" target="_blank"><%=rs("title")%></a></td>
            <td><%= FormatDateTime(rs("UpdateTime"),2) %></td>
            <td><%=rs("Hits")%>/<%=rs("Votes")%></td>
            <td> 
              <%if rs("Passed")=true then%>
              已审核 
              <%else%>
              <font color="#FF0000">未审核</font> 
              <%end if%>
              <%if rs("showMov")=true then%>
              隐藏 
              <%else%>
              <font color="#FF0000">显示</font> 
              <%end if%>
            </td>
            <td> 
              <a href="Bs_Article_edit.asp?ArticleID=<%=rs("articleid")%>">修改</a> 
              <a href="Bs_Article_Del.asp?ArticleID=<%=rs("ArticleID")%>&Action=Del" onClick="return ConfirmDel();">删除</a> 
            </td>
          </tr>
          <%
	i=i+1
	      if i>=MaxPerPage then exit do
	      rs.movenext
	loop
%>
        </table>
        <table width="100%">
          <tr> 
            <td><input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox">
              选中本页显示的所有</td>
            <td><input name="submit" type='submit' value='删除选定的项'> 
              <input name="Action" type="hidden" id="Action" value="Del"></td>
          </tr>
        </table>
        <%
   end sub 
%>
      </form>
      <br> 
      <table width="100%">
        <tr>
          <form name="searchsoft" method="get" action="Bs_Article.asp">
            <td height="25"> <strong>查找：</strong> 
              <input name="Title" type="text" class=smallInput id="Title3" size="20" maxlength="50"> 
              <input name="Query" type="submit" id="Query" value="查 询">
            请输入名称。如果为空，则查找所有。</td>
          </form>
        </tr>
      </table>
		</td>
	</tr>
</table>
<%htmlend%>
<%
rs.close
set rs=nothing  
call CloseConn()
%>