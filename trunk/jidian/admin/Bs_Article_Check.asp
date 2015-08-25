<!--#include file="bsconfig.asp"-->
<!--#include file="../Inc/Ubbcode.asp"-->
<!--#include file="Inc/Function.asp"-->

<%
dim strFileName
const MaxPerPage=20
dim totalPut   
dim CurrentPage
dim TotalPages
dim i,j
dim ArticleID
dim Title
dim BigClassName,SmallClassName,SpecialName
dim Passed
dim PurviewChecked
dim strAdmin,arrAdmin

PurviewChecked=false
Passed=trim(request("Passed"))
strFileName="Bs_Article_Check.asp?BigClassName=" & BigClassName & "&SmallClassName=" & SmallClassName & "&SpecialName=" & SpecialName

if Passed="" then
	if Session("Passed")="" then
		Passed="False"
	else
		Passed=Session("Passed")
	end if
end if
session("Passed")=Passed
Title=Trim(request("Title"))
ArticleID=Request("ArticleID")
BigClassName=Trim(request("BigClassName"))
SmallClassName=Trim(request("SmallClassName"))
if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if

dim sql
dim rs
sql="select * from Product where Passed=" & Passed
if Title<>"" then
	sql=sql & " and title like '%" & Title & "%' "
else
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
end if
sql=sql & " order by articleid desc"
Set rs= Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,1
%>
<SCRIPT language=javascript>
function unselectall()
{
    if(document.Check.chkAll.checked){
	document.Check.chkAll.checked = document.Check.chkAll.checked&0;
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
</SCRIPT>
<!-- #include file="Inc/Head.asp" -->
<%call checkmanage("53")%>
<table align="center" class="a2">
	<tr>
		<td class="a1">网络课程审核</td>
	</tr>
	<tr class="a4">
		<td>
      <form name="Passed" method="Get" action="Bs_Article_Check.asp">
				<div align="center"><strong>选项：</strong> 
				<input name="Passed" type="radio" value="False" onClick="submit();" <%if Passed="False" then response.write " checked"%>>
				未审核的&nbsp;&nbsp;&nbsp;&nbsp; 
				<input name="Passed" type="radio" value="True" onClick="submit();" <%if Passed="True" then response.write " checked"%>>
				已审核的
				<input name="BigClassName" type="hidden" id="BigClassName" value="<%=BigClassName%>">
				<input name="SmallClassName" type="hidden" id="SmallClassName" value="<%=SmallClassName%>">
				<input name="SpecialName" type="hidden" id="SpecialName" value="<%=SpecialName%>">
		</div>
		  </form>
			<table width="100%">
				<tr> 
					<td>|&nbsp; 
<%
dim sqlBigClass,sqlSmallClass,rsBigClass,rsSmallClass,sqlSpecial,rsSpecial
sqlBigClass="select * from BigClass"
Set rsBigClass= Server.CreateObject("ADODB.Recordset")
rsBigClass.open sqlBigClass,conn,1,1
if rsBigClass.eof then 
	response.Write("还没有任何类别，请首先添加类别。")
end if
do while not rsBigClass.eof
	if rsBigClass("BigClassName")=BigClassName then
		response.Write("<a href='Bs_Article_Check.asp?BigClassName=" & rsBigClass("BigClassName") & "'><font color='red'>" & rsBigClass("BigClassName") & "</font></a> | ")
		if session("purview")=3 then
			strAdmin=rsBigClass("Admin")
			if Instr(strAdmin,"|")>0 then
				arrAdmin=split(strAdmin)
				for i=0 to ubound(arrAdmin)
					if trim(arrAdmin(i))=session("admin") then
						PurviewChecked=True
						exit for
					end if
				next
			else
				if trim(strAdmin)=session("Admin") then
					PurviewChecked=True
				end if
			end if
		end if
	else
		response.Write("<a href='Bs_Article_Check.asp?BigClassName=" & rsBigClass("BigClassName") & "'>" & rsBigClass("BigClassName") & "</a> | ")
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
		response.write "<tr class='tdbg'><td bgcolor='#CCCCCC'>"
		do while not rsSmallClass.eof
			if rsSmallClass("SmallClassName")=SmallClassName then
				response.Write("&nbsp;<a href='Bs_Article_Check.asp?BigClassName=" & rsSmallClass("BigClassName") & "&SmallClassName=" & rsSmallClass("SmallClassName") & "'><font color='red'>" & rsSmallClass("SmallClassName") & "</font></a>&nbsp;&nbsp;")
				if session("purview")=4 then
					strAdmin=rsSmallClass("Admin")
					if Instr(strAdmin,"|")>0 then
						arrAdmin=split(strAdmin)
						for i=0 to ubound(arrAdmin)
							if trim(arrAdmin(i))=session("admin") then
								PurviewChecked=True
								exit for
							end if
						next
					else
						if trim(strAdmin)=session("Admin") then
							PurviewChecked=True
						end if
					end if
				end if
			else
				response.Write("&nbsp;<a href='Bs_Article_Check.asp?BigClassName=" & rsSmallClass("BigClassName") & "&SmallClassName=" & rsSmallClass("SmallClassName") & "'>" & rsSmallClass("SmallClassName") & "</a>&nbsp;&nbsp;")
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
			<form action="Bs_Article_Check_Set.asp" method="Post" name="Check" id="Check">
			<table width="100%">
				<tr> 
					<td width="453" height="22" bgcolor="#CCCCCC"><a href="Bs_Article_Check.asp">&nbsp;审核</a>&nbsp;&gt;&gt;&nbsp; 
<%
if Title="" and BigClassName="" and SmallClassName="" and SpecialName="" then
	if Passed="False" then
		response.write "所有未审核"
	else
		response.write "所有已审核"
	end if
else
	if request("Query")<>"" then
		if Title<>"" then
			response.write "标题中含有“<font color=blue>" & Title & "</font>”并且未审核的文章"
		else
			response.Write("所有未审核的文章")
		end if
 	else
		if BigClassName<>"" then
			response.write "<a href='Bs_Article_Check.asp?BigClassName=" & BigClassName & "'>" & BigClassName & "</a>&nbsp;&gt;&gt;&nbsp;"
			if SmallClassName<>"" then
				response.write "<a href='Bs_Article_Check.asp?BigClassName=" & BigClassName & "&SmallClassName=" & SmallClassName & "'>" & SmallClassName & "</a>"
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
					<td width="161" bgcolor="#CCCCCC"> &nbsp; 
<%
	if rs.eof and rs.bof then
	response.write "共找到 0 个</td></tr></table>"
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
	response.Write "共找到 " & totalPut & " 个"
%>
					</td>
				</tr>
			</table>
<%		
		if currentPage=1 then
				showContent
				showpage strFileName,totalput,MaxPerPage,true,false," 个"
		else
				if (currentPage-1)*MaxPerPage<totalPut then
						rs.move  (currentPage-1)*MaxPerPage
					dim bookmark
						bookmark=rs.bookmark
						showContent
						showpage strFileName,totalput,MaxPerPage,true,false," 个"
				else
					currentPage=1
						showContent
						showpage strFileName,totalput,MaxPerPage,true,false," 个"
			end if
	end if
end if
%>
<%  
sub showContent
   	dim i
    i=0
%>
			<table width="100%" style="word-break:break-all">
				<tr class="opttitle"> 
					<td>选择</td>
					<td>ID</td>
					<td>名称</td>
					<td>作者</td>
					<td>加入时间</td>
					<td>点击数</td>
					<td>已审核</td>
					<td>操作</td>
				</tr>
<%do while not rs.eof%>
				<tr class="opt"> 
					<td> 
					<input name='ArticleID' type='checkbox' onClick="unselectall()" id="ArticleID" value='<%=cstr(rs("articleID"))%>'>
					</td>
					<td><%=rs("articleid")%></td>
					<td> <a href="../ArticleShow.asp?ArticleID=<%=rs("articleid")%>" 
					title="<%=replace(left(nohtml(rs("Content")),200),chr(34),"") & "……"%>">&nbsp;<%=rs("title")%></a></td>
					<td><%=rs("author")%></td>
					<td><%= FormatDateTime(rs("UpdateTime"),2) %></td>
					<td><%=rs("Hits")%> </td>
					<td>
<%if rs("Passed")=true then response.write "是" else response.write "否" end if%>
					</td>
					<td> 
<%
if session("purview")<=2 or PurviewChecked=True then
	If rs("Passed")=False Then 
%>
<a href="Bs_Article_Check_Set.asp?Action=Check&ArticleID=<%=rs("ArticleID")%>">通过审核</a> 
<% Else %>
<a href="Bs_Article_Check_Set.asp?Action=CancelCheck&ArticleID=<%=rs("ArticleID")%>">取消通过</a> 
<%
	End If
end if
%>
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
				  <td width="250" height="30">
					<input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox">					选中本页显示的所有</td>
					<td><input name="submit" type='submit'
					value="<%if Passed="True" then response.write "取消"%>审核选定" 
					<%if  session("purview")=4 and PurviewChecked=False then response.write "disabled"%>> 
					<input name="Action" type="hidden" id="Action" value="
<%
if Passed="False" then
response.write "Check"
else
response.write "CancelCheck"
end if%>"></td>
				</tr>
			</table>
<%
end sub 
%>
			</form>
			<table width="100%">
				<tr> 
					<form name="searchsoft" method="get" action="Bs_Article_Check.asp">
					<td height="30"> <strong>查找：</strong> 
					<input name="Title" type="text" class=smallInput id="Title3" size="20"> 
					<input name="Query" type="submit" id="Query" value="查 询">
					&nbsp;&nbsp;请输入名称。如果为空，则查找所有。 </td>
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

sub CheckArticle(id)
    dim sql
    sql="update article set Passed = True where articleid="&cstr(id)
    conn.execute sql
    if err.Number<>0 then
		err.clear
		'response.write "删 除 失 败 !<br>"
    else
        'response.write "删除成功！<br>"
    end if	
End sub
sub CancelCheckArticle(id)
    dim sql
    sql="update article set Passed = False where articleid="&cstr(id)
    conn.execute sql
    if err.Number<>0 then
		err.clear
		'response.write "删 除 失 败 !<br>"
    else
        'response.write "删除成功！<br>"
    end if	
End sub
%>