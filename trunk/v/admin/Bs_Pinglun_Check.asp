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
dim PinglunID
dim shenhe
dim strAdmin,arrAdmin

shenhe=trim(request("shenhe"))
strFileName="Bs_Pinglun_Check.asp?id=" & id 

if shenhe="" then
	if Session("shenhe")="" then
		shenhe="False"
	else
		shenhe=Session("shenhe")
	end if
end if
session("shenhe")=shenhe
name=Trim(request("name"))
PinglunID=Request("PinglunID")
if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if

dim sql
dim rs
sql="select * from buyok_pinglun where shenhe=" & shenhe
if name<>"" then
	sql=sql & " and name like '%" & name & "%' "
else
	sql=sql
end if
sql=sql & " order by id desc"
Set rs= Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,1
%>
<SCRIPT language=javascript>
function unselectall()
{
    if(document.CheckB.chkAll.checked){
	document.CheckB.chkAll.checked = document.CheckB.chkAll.checked&0;
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
  
function shenHe() {
	    document.forms[1].target="_self";
		document.forms[1].action="Bs_Pinglun_Check_Set.asp";
	    document.forms[1].submit();
}
function shanChu() {
	    document.forms[1].target="_self";
		document.forms[1].action="Bs_Pinglun_Del.asp";
	    document.forms[1].submit();
}

function ConfirmDel()
{
   if(confirm("ȷ��Ҫɾ��ѡ�еĲ�Ʒ��һ��ɾ�������ָܻ���"))
     return true;
   else
     return false;
	 
}
</SCRIPT>
<!-- #include file="Inc/Head.asp" -->
<%call checkmanage("32")%>
<table align="center" class="a2">
	<tr>
		<td class="a1">�� �� �� ��</td>
	</tr>
	<tr class="a4">
		<td>
      <form name="shenhe" method="Get" action="Bs_Pinglun_Check.asp">
				<div align="center"><strong>ѡ�</strong> 
				<input name="shenhe" type="radio" value="False" onClick="submit();" <%if shenhe="False" then response.write " checked"%>>
				δ��˵�&nbsp;&nbsp;&nbsp;&nbsp; 
				<input name="shenhe" type="radio" value="True" onClick="submit();" <%if shenhe="True" then response.write " checked"%>>
				����˵�
				<input name="name" type="hidden" id="name" value="<%=name%>">
				<input name="prodid" type="hidden" id="prodid" value="<%=prodid%>">
		</div>
		  </form>

			<form  method="Post" name="CheckB" id="CheckB">
			<table width="100%">
				<tr> 
					<td width="453" height="22" bgcolor="#CCCCCC"><a href="Bs_Pinglun_Check.asp">&nbsp;���</a>&nbsp;&gt;&gt;&nbsp; 
<%
if name="" then
	if shenhe="False" then
		response.write "����δ��˵���Ƶ"
	else
		response.write "��������˵���Ƶ"
	end if
else
	if request("Query")<>"" then
		if name<>"" then
			response.write "�û����к��С�<font color=blue>" & name & "</font>������δ��˵���Ƶ"
		else
			response.Write("����δ��˵���Ƶ")
		end if
 	else
		if name<>"" then
			response.write "<a href='Bs_Pinglun_Check.asp?id=" & id & "'>" & name & "</a>&nbsp;&gt;&gt;&nbsp;"
		end if
	end if
end if
%>
				  </td>
					<td width="161" bgcolor="#CCCCCC"> &nbsp; 
<%
	if rs.eof and rs.bof then
	response.write "���ҵ� 0 ����Ƶ</td></tr></table>"
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
	response.Write "���ҵ� " & totalPut & " ����Ƶ"
%>
					</td>
				</tr>
			</table>
<%		
		if currentPage=1 then
				showContent
				showpage strFileName,totalput,MaxPerPage,true,false," ����Ƶ"
		else
				if (currentPage-1)*MaxPerPage<totalPut then
						rs.move  (currentPage-1)*MaxPerPage
					dim bookmark
						bookmark=rs.bookmark
						showContent
						showpage strFileName,totalput,MaxPerPage,true,false," ����Ƶ"
				else
					currentPage=1
						showContent
						showpage strFileName,totalput,MaxPerPage,true,false," ����Ƶ"
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
					<td>��������</td>
					<td>����ʱ��</td>
					<td>�û���</td>
					<td>�����</td>
					<td>����</td>
				</tr>
<%do while not rs.eof%>
				<tr class="opt"> 
					<td> 
					<input name='PinglunID' type='checkbox' onClick="unselectall();" id="PinglunID" value='<%=cstr(rs("id"))%>'>
                    <input name='prodid' type='hidden' id="prodid" value='<%=rs("prodid")%>'>
					</td>
					<td><%=rs("id")%></td>
					<td> <a href="../ArticleShow.asp?ArticleID=<%=rs("prodid")%>" 
					title="<%=rs("nr")%>"> <%=replace(left(nohtml(rs("nr")),200),chr(34),"") & "����"%></a></td>
					<td><%= FormatDateTime(rs("adddate"),2) %></td>
					<td><%=rs("name")%> </td>
					<td>
<%if rs("shenhe")=true then response.write "��" else response.write "��" end if%>
					</td>
					<td> 
<%
	If rs("shenhe")=False Then 
%>
<a href="Bs_Pinglun_Check_Set.asp?ActionH=OkCheck&PinglunID=<%=rs("id")%>">ͨ�����</a> 
<% Else %>
<a href="Bs_Pinglun_Check_Set.asp?ActionH=CancelCheck&PinglunID=<%=rs("id")%>">ȡ��ͨ��</a> 
<%
	End If
%>
<a href="Bs_Pinglun_Del.asp?PinglunID=<%=rs("id")%>&Act=Del" onClick="return ConfirmDel();">ɾ��</a> 
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
					<input name="chkAll" type="checkbox" id="chkAll" onclick="CheckAll(this.form)" value="checkbox">					ѡ�б�ҳ��ʾ������</td>
					<td>
                    
                    <input name="shanchu" type="button" value="ɾ��ѡ������" onclick="shanChu()"> 
                    <input name="Act" type="hidden" id="Act" value="Del">
                    <input name="shenhe" type='button' value="<%if shenhe="True" then response.write "ȡ��"%>���ѡ��"  <%'response.write "disabled"%> onclick='shenHe()'> 
					<input name="ActionH" type="hidden" id="ActionH" value="<%if shenhe="False" then
response.write "OkCheck"
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
					<form name="searchsoft" method="get" action="Bs_Pinglun_Check.asp">
					<td height="30"> <strong>���ң�</strong> 
					<input name="name" type="text" class=smallInput id="name" size="20"> 
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

%>