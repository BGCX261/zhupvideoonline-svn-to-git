<!--#include file="bsconfig.asp"-->
<!--#include file="Inc/Function.asp"-->

<%
dim strFileName
const MaxPerPage=20
dim totalPut,CurrentPage,TotalPages
dim rs, sql
strFileName="Bs_User.asp"

if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if

Set rs=Server.CreateObject("Adodb.RecordSet")
sql="select * from [User] order by UserID desc"
rs.Open sql,conn,1,1
%>
<script language=javascript>
function ConfirmDel()
{
   if(confirm("ȷ��Ҫɾ�����û���"))
     return true;
   else
     return false;
}
</script>
<!-- #include file="Inc/Head.asp" -->
<%call checkmanage("31")%>
<table  class="bsOutTable">
	<tr>
		<td class="bsNavTitle">ע���Ա����</td>
	</tr>
	<tr class="bsMainTd">
		<td>
<%
  	if rs.eof and rs.bof then
		response.write "Ŀǰ���� 0 ��ע���û�"
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
	    if currentPage=1 then
        	showpage strFileName,totalput,MaxPerPage,true,true,"���û�"
        	showContent
        	showpage strFileName,totalput,MaxPerPage,true,true,"���û�"
   	 	else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rs.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rs.bookmark
        		showpage strFileName,totalput,MaxPerPage,true,true,"���û�"
            	showContent
            	showpage strFileName,totalput,MaxPerPage,true,true,"���û�"
        	else
	        	currentPage=1
        		showpage strFileName,totalput,MaxPerPage,true,true,"���û�"
           		showContent
           		showpage strFileName,totalput,MaxPerPage,true,true,"���û�"
	    	end if
		end if
	end if

sub showContent
   	dim i
    i=0
%>
      <table  class="bsIptTable">
        <tr class="optTitle"> 
          <td>���</td>
          <td>�û���</td>
          <td>�Ա�</td>
          <td>Email</td>
          <td>��ʵ����</td>
          <td>״̬</td>
          <td>����</td>
        </tr>
        <%do while not rs.EOF %>
        <tr class="optCon"> 
          <td><%=rs("UserID")%></td>
          <td><%=rs("UserName")%></td>
          <td> 
            <%if rs("Sex")=1 then
	    response.write "��"
	  else
	    response.write "Ů"
	  end if%>
          </td>
          <td><%=rs("email")%></td>
          <td> <%=rs("Comane")%> </td>
          <td> 
            <%
	  if rs("LockUser")=true then
	  	response.write "<font color='#ff0000'>δ���</font>"
	  else
	  	response.write "<font color='#0000ff'>�����</font>"
	  end if
	  %>
          </td>
          <td>
          <a href="Bs_User_edit.asp?UserID=<%=rs("UserID")%>">�޸�</a>&nbsp; 
            <%if rs("LockUser")=False then %>
            <a href="Bs_User_Lock.asp?Action=Lock&UserID=<%=rs("UserID")%>">����</a> 
            <%else%>
            <a href="Bs_User_Lock.asp?Action=CancelLock&UserID=<%=rs("UserID")%>">���ͨ��</a> 
            <%end if%>
            &nbsp;<a href="Bs_User_Del.asp?UserID=<%=rs("UserID")%>" onClick="return ConfirmDel();">ɾ��</a></td>
        </tr>
        <%
	i=i+1
	if i>=MaxPerPage then exit do
	rs.movenext
loop
%>
      </table>
      <%
end sub 
%></td>
	</tr>
</table>
<%htmlend%>
<%
rs.Close
set rs=Nothing
call CloseConn()  
%>