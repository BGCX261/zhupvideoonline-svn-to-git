<!--#include file="security.asp"-->
<!--#include file="dbConn.asp"-->
<!--#include file="head.asp"-->
<%
sql="select * from baseSet"
Set rs= Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,1
   dim rs
   MaxPerPage=rs("duo")
   dim totalPut   
   dim CurrentPage
   dim TotalPages
   dim i,j
   dim typename
   typename=""
   if not isempty(request("page")) then
      currentPage=cint(request("page"))
   else
      currentPage=1
   end if
   dim sql
   dim rstype
   dim typesql
   dim typeid,typename1
   if not isEmpty(request("typeid")) then
	typeid=request("typeid")
   else
	typeid=0
   end if
 set rstype=server.createobject("adodb.recordset")
  typesql="select * from type where typeid="&typeid&""
 rstype.open typesql,conn,1,1
 if not rstype.eof then
	typename1=rstype("type")
 else typename1="全部写真"	
 end if
rstype.close
   set rst=server.CreateObject("ADODB.RecordSet")
  
%>
<body>
<div class="admCont"> 
  <div class="admTop"></div>
  <div class="admLeft">
      <!--#include file="left.asp"-->
  </div>
  <div class="admRight">
      <div class="admPhotoType">
          <ul>
            <%
            rst.open "select * from type",conn,1
         if rst.EOF then
            response.write "没有栏目："
            else
                %>  <li><a href="managePhoto.asp?typeid=0">全部图片</a></li> 
                  <%do while NOT rst.EOF%>
                  <li><a href="managePhoto.asp?typeid=<%=rst("typeid")%>"><%=rst("type")%></a></li>
                  <%
            rst.MoveNext
            loop
            end if
        rst.close
                  %>
          </ul>
      </div>
  <% 
if typeid=0 then   
 sql="select articleid,title,dateandtime from albums  order by dateandtime desc" 
else   
 sql="select articleid,title,dateandtime from albums where typeid="+cstr(typeid)+" order by dateandtime desc" 
end if 
Set rs= Server.CreateObject("ADODB.Recordset") 
rs.open sql,conn,1,1 
  if rs.eof and rs.bof then 
       response.write "<p>暂无数据！</p>" 
   else 
	  totalPut=rs.recordcount 
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
           showpage totalput,MaxPerPage,"managePhoto.asp" 
            showContent 
            showpage totalput,MaxPerPage,"managePhoto.asp" 
       else 
          if (currentPage-1)*MaxPerPage<totalPut then 
            rs.move  (currentPage-1)*MaxPerPage 
            dim bookmark 
            bookmark=rs.bookmark 
           showpage totalput,MaxPerPage,"managePhoto.asp" 
            showContent 
             showpage totalput,MaxPerPage,"managePhoto.asp" 
        else 
	        currentPage=1 
           showpage totalput,MaxPerPage,"managePhoto.asp" 
           showContent 
           showpage totalput,MaxPerPage,"managePhoto.asp" 
	      end if 
	   end if 
   end if  
 
   sub showContent 
       dim i 
	   i=0 
 
  %>
    <table class="admListTable">
      <tr class="admListTitle"> 
        <td width="64"> ID号</td>
        <td width="400"> 图片标题</td>
        <td width="68"> 修改</td>
        <td width="68"> 删除</td>
      </tr>
      <%do while not rs.eof%>
      <tr class="admListDetail"> 
        <td><%=rs("articleid")%> </td>
        <td><a href="photoDetail.asp?id=<%=rs("articleid")%>" target="_blank"><%=rs("title")%></a></td>
        <td><a href="editPhoto.asp?id=<%=rs("articleid")%>">修改</a></td>
        <td><a href="deletePhoto.asp?articleid=<%=rs("articleid")%>&page=<%=request("page")%>&typeid=<%=request("typeid")%>"> 删除</a></td>
      </tr>
      <% i=i+1 
	      if i>=MaxPerPage then exit do 
	      rs.movenext 
	   loop 
		  %>
    </table>
  <% 
   end sub  
 
function showpage(totalnumber,maxperpage,filename) 
  dim n 
  if totalnumber mod maxperpage=0 then 
     n= totalnumber \ maxperpage 
  else 
     n= totalnumber \ maxperpage+1 
  end if 
  response.write "<div class=pInfo><form method=Post action="&filename&"?typeid="&typeid&">" 
    if CurrentPage<2 then 
    response.write "<font color='#000080'>首页 上一页</font>&nbsp;" 
  else 
    response.write "<a href="&filename&"?page=1&typeid="&typeid&">首页</a>&nbsp;" 
    response.write "<a href="&filename&"?page="&CurrentPage-1&"&typeid="&typeid&">上一页</a>&nbsp;" 
  end if 
  if n-currentpage<1 then 
    response.write "<font color='#000080'>下一页 尾页</font>" 
  else 
    response.write "<a href="&filename&"?page="&(CurrentPage+1)&"&typeid="&typeid&">" 
    response.write "下一页</a> <a href="&filename&"?page="&n&"&typeid="&typeid&">尾页</a>" 
  end if 
   response.write "<font color='#000080'>&nbsp;页次：</font><font color=red>"&CurrentPage&"</font><font color='#000080'>/"&n&"页</font> " 
    response.write "<font color='#000080'>&nbsp;共<b>"&totalnumber&"</b>篇 <b>"&maxperpage&"</b>篇/页</font> " 
   response.write " <font color='#000080'>转到：</font><input type='text' name='page' size=4 maxlength=10 value="&currentpage&">" 
   response.write "<input type='submit'  value='Go'  name='cndok'></span></p></form></div>" 
        
end function 
  %>
    </div>
</div>
<%rs.close          
set rs=nothing   
conn.close 
set conn=nothing %>
</body>
</html>


