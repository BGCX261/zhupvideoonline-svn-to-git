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
		<td class="a1">����γ����</td>
	</tr>
	<tr class="a4">
		<td>
      <form name="Passed" method="Get" action="Bs_Article_Check.asp">
				<div align="center"><strong>ѡ�</strong> 
				<input name="Passed" type="radio" value="False" onClick="submit();" <%if Passed="False" then response.write " checked"%>>
				δ��˵�&nbsp;&nbsp;&nbsp;&nbsp; 
				<input name="Passed" type="radio" value="True" onClick="submit();" <%if Passed="True" then response.write " checked"%>>
				����˵�
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
	response.Write("��û���κ����������������")
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
					<td width="453" height="22" bgcolor="#CCCCCC"><a href="Bs_Article_Check.asp">&nbsp;���</a>&nbsp;&gt;&gt;&nbsp; 
<%
if Title="" and BigClassName="" and SmallClassName="" and SpecialName="" then
	if Passed="False" then
		response.write "����δ���"
	else
		response.write "���������"
	end if
else
	if request("Query")<>"" then
		if Title<>"" then
			response.write "�����к��С�<font color=blue>" & Title & "</font>������δ��˵�����"
		else
			response.Write("����δ��˵�����")
		end if
 	else
		if BigClassName<>"" then
			response.write "<a href='Bs_Article_Check.asp?BigClassName=" & BigClassName & "'>" & BigClassName & "</a>&nbsp;&gt;&gt;&nbsp;"
			if SmallClassName<>"" then
				response.write "<a href='Bs_Article_Check.asp?BigClassName=" & BigClassName & "&SmallClassName=" & SmallClassName & "'>" & SmallClassName & "</a>"
			else
				response.write "����С��"
			end if
		end if
		if SpecialName<>"" then
			response.write "<font color=red>[ר��]</font> " & SpecialName
		end if
	end if
end if
%>
				  </td>
					<td width="161" bgcolor="#CCCCCC"> &nbsp; 
<%
	if rs.eof and rs.bof then
	response.write "���ҵ� 0 ��</td></tr></table>"
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
	response.Write "���ҵ� " & totalPut & " ��"
%>
					</td>
				</tr>
			</table>
<%		
		if currentPage=1 then
				showContent
				showpage strFileName,totalput,MaxPerPage,true,false," ��"
		else
				if (currentPage-1)*MaxPerPage<totalPut then
						rs.move  (currentPage-1)*MaxPerPage
					dim bookmark
						bookmark=rs.bookmark
						showContent
						showpage strFileName,totalput,MaxPerPage,true,false," ��"
				else
					currentPage=1
						showContent
						showpage strFileName,totalput,MaxPerPage,true,false," ��"
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
					<td>ѡ��</td>
					<td>ID</td>
					<td>����</td>
					<td>����</td>
					<td>����ʱ��</td>
					<td>�����</td>
					<td>�����</td>
					<td>����</td>
				</tr>
<%do while not rs.eof%>
				<tr class="opt"> 
					<td> 
					<input name='ArticleID' type='checkbox' onClick="unselectall()" id="ArticleID" value='<%=cstr(rs("articleID"))%>'>
					</td>
					<td><%=rs("articleid")%></td>
					<td> <a href="../ArticleShow.asp?ArticleID=<%=rs("articleid")%>" 
					title="<%=replace(left(nohtml(rs("Content")),200),chr(34),"") & "����"%>">&nbsp;<%=rs("title")%></a></td>
					<td><%=rs("author")%></td>
					<td><%= FormatDateTime(rs("UpdateTime"),2) %></td>
					<td><%=rs("Hits")%> </td>
					<td>
<%if rs("Passed")=true then response.write "��" else response.write "��" end if%>
					</td>
					<td> 
<%
if session("purview")<=2 or PurviewChecked=True then
	If rs("Passed")=False Then 
%>
<a href="Bs_Article_Check_Set.asp?Action=Check&ArticleID=<%=rs("ArticleID")%>">ͨ�����</a> 
<% Else %>
<a href="Bs_Article_Check_Set.asp?Action=CancelCheck&ArticleID=<%=rs("ArticleID")%>">ȡ��ͨ��</a> 
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
					<input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox">					ѡ�б�ҳ��ʾ������</td>
					<td><input name="submit" type='submit'
					value="<%if Passed="True" then response.write "ȡ��"%>���ѡ��" 
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
					<td height="30"> <strong>���ң�</strong> 
					<input name="Title" type="text" class=smallInput id="Title3" size="20"> 
					<input name="Query" type="submit" id="Query" value="�� ѯ">
					&nbsp;&nbsp;���������ơ����Ϊ�գ���������С� </td>
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
		'response.write "ɾ �� ʧ �� !<br>"
    else
        'response.write "ɾ���ɹ���<br>"
    end if	
End sub
sub CancelCheckArticle(id)
    dim sql
    sql="update article set Passed = False where articleid="&cstr(id)
    conn.execute sql
    if err.Number<>0 then
		err.clear
		'response.write "ɾ �� ʧ �� !<br>"
    else
        'response.write "ɾ���ɹ���<br>"
    end if	
End sub
%>