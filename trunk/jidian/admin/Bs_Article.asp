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
dim PurviewChecked
dim strAdmin,arrAdmin

PurviewChecked=false
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
if session("purview")>4 then
	sql=sql & " and Editor='" & Session("admin") & "' and Passed=false"
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
   if(confirm("ȷ��Ҫɾ��ѡ�е�����һ��ɾ�������ָܻ���"))
     return true;
   else
     return false;
	 
}

</SCRIPT>
<!-- #include file="Inc/Head.asp" -->
<%call checkmanage("49")%>
<table align="center" class="a2">
	<tr>
		<td class="a1">����γ̹���</td>
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
	response.Write("��û���κ���Ŀ�������������Ŀ��")
end if
do while not rsBigClass.eof
	if Server.URLEncode(rsBigClass("BigClassName"))=BigClassName then
		response.Write("<a href='Bs_Article.asp?BigClassName=" & Server.URLEncode(rsBigClass("BigClassName")) & "'><font color='red'>" & rsBigClass("BigClassName") & "</font></a> | ")
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
            <td height="25" bgcolor="#CCCCCC"><a href="Bs_Article.asp">����γ̹���</a> 
              &gt;&gt; 
              <%
if request.querystring="" then
	response.write "��������γ�"
else
	if request("Query")<>"" then
		if Title<>"" then
			response.write "�����к��С�<font color=blue>" & Title & "</font>��������γ�"
		else
			response.Write("��������γ�")
		end if
 	else
		if BigClassName<>"" then
			response.write "<a href='Bs_Article.asp?BigClassName=" & Server.URLEncode(BigClassName) & "'>" & BigClassName & "</a>&nbsp;&gt;&gt;&nbsp;"
			if SmallClassName<>"" then
				response.write "<a href='Bs_Article.asp?BigClassName=" & Server.URLEncode(BigClassName) & "&SmallClassName=" & Server.URLEncode(SmallClassName) & "'>" & SmallClassName & "</a>"
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
            <td width="150" bgcolor="#CCCCCC">&nbsp;
            <%
  	if rs.eof and rs.bof then
		response.write "���ҵ� 0 ����Ϣ</td></tr></table>"
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
		response.Write "���ҵ� " & totalPut & " ����Ϣ"
%>
            </td>
          </tr>
        </table>
        <%		
	    if currentPage=1 then
        	showContent
        	showpage strFileName,totalput,MaxPerPage,true,false,"����Ϣ"
   	 	else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rs.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rs.bookmark
            	showContent
            	showpage strFileName,totalput,MaxPerPage,true,false,"����Ϣ"
        	else
	        	currentPage=1
           		showContent
           		showpage strFileName,totalput,MaxPerPage,true,false,"����Ϣ"
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
            <td>ѡ��</td>
            <td>ID</td>
            <td>���</td>
            <td>����</td>
            <td>����</td>
            <td>����ʱ��</td>
            <td>������</td>
            <td>����</td>
          </tr>
          <%do while not rs.eof%>
          <tr class="opt"> 
            <td> <input name='ArticleID' type='checkbox' onClick="unselectall()" id="ArticleID" value='<%=cstr(rs("articleID"))%>'> 
            </td>
            <td><%=rs("articleid")%></td>
            <td><%=rs("Product_Id")%></td>
            <td><a href="../ArticleShow.asp?ArticleID=<%=rs("articleid")%>" target="_blank"><%=rs("title")%></a></td>
            <td><%=rs("author")%></td>
            <td><%= FormatDateTime(rs("UpdateTime"),2) %></td>
            <td> 
              <%if rs("Passed")=true then%>
              ����� 
              <%else%>
              <font color="#FF0000">δ���</font> 
              <%end if%>
            </td>
            <td> 
              <a href="Bs_Article_edit.asp?ArticleID=<%=rs("articleid")%>">�޸�</a> 
              <%if session("purview")=4 then%><%else%> <a href="Bs_Article_Del.asp?ArticleID=<%=rs("ArticleID")%>&Action=Del" onClick="return ConfirmDel();">ɾ��</a><%end if%>
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
              ѡ�б�ҳ��ʾ������</td>
            <td><input name="submit" type='submit' value='ɾ��ѡ������' <%if session("purview")=4 and PurviewChecked=False then response.write "disabled"%>> 
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
            <td height="25"> <strong>���ң�</strong> 
              <input name="Title" type="text" class=smallInput id="Title3" size="20" maxlength="50"> 
              <input name="Query" type="submit" id="Query" value="�� ѯ">
            ���������ơ����Ϊ�գ���������С�</td>
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