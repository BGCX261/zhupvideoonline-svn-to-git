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
   if(confirm("确定要删除此用户吗？"))
     return true;
   else
     return false;
}
</script>
<!-- #include file="Inc/Head.asp" -->
<%call checkmanage("31")%>
<table  class="bsOutTable">
	<tr>
		<td class="bsNavTitle">注册会员管理</td>
	</tr>
	<tr class="bsMainTd">
		<td>
<%
  	if rs.eof and rs.bof then
		response.write "目前共有 0 个注册用户"
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
        	showpage strFileName,totalput,MaxPerPage,true,true,"个用户"
        	showContent
        	showpage strFileName,totalput,MaxPerPage,true,true,"个用户"
   	 	else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rs.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rs.bookmark
        		showpage strFileName,totalput,MaxPerPage,true,true,"个用户"
            	showContent
            	showpage strFileName,totalput,MaxPerPage,true,true,"个用户"
        	else
	        	currentPage=1
        		showpage strFileName,totalput,MaxPerPage,true,true,"个用户"
           		showContent
           		showpage strFileName,totalput,MaxPerPage,true,true,"个用户"
	    	end if
		end if
	end if

sub showContent
   	dim i
    i=0
%>
      <table  class="bsIptTable">
        <tr class="optTitle"> 
          <td>序号</td>
          <td>用户名</td>
          <td>性别</td>
          <td>Email</td>
          <td>真实名称</td>
          <td>状态</td>
          <td>操作</td>
        </tr>
        <%do while not rs.EOF %>
        <tr class="optCon"> 
          <td><%=rs("UserID")%></td>
          <td><%=rs("UserName")%></td>
          <td> 
            <%if rs("Sex")=1 then
	    response.write "男"
	  else
	    response.write "女"
	  end if%>
          </td>
          <td><%=rs("email")%></td>
          <td> <%=rs("Comane")%> </td>
          <td> 
            <%
	  if rs("LockUser")=true then
	  	response.write "<font color='#ff0000'>未审核</font>"
	  else
	  	response.write "<font color='#0000ff'>已审核</font>"
	  end if
	  %>
          </td>
          <td>
          <a href="Bs_User_edit.asp?UserID=<%=rs("UserID")%>">修改</a>&nbsp; 
            <%if rs("LockUser")=False then %>
            <a href="Bs_User_Lock.asp?Action=Lock&UserID=<%=rs("UserID")%>">锁定</a> 
            <%else%>
            <a href="Bs_User_Lock.asp?Action=CancelLock&UserID=<%=rs("UserID")%>">审核通过</a> 
            <%end if%>
            &nbsp;<a href="Bs_User_Del.asp?UserID=<%=rs("UserID")%>" onClick="return ConfirmDel();">删除</a></td>
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