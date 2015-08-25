<!--#include file="dbConn.asp"-->
<!--#include file="head.asp"-->
<%
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
 else typename1="全部图集"	
 end if
rstype.close
'---------------------------search----------------------
dim keyword
keyword=trim(request("keyword"))
keyword=replace(keyword,"'","''")
if typeid=0 then
    sql="select * from albums where title Like '%"& keyword &"%' or title like title Like '%"& keyword &"%' or title Like '%"& keyword &"%'order by articleid desc"
else
   sql="select * from albums where (typeid="&typeid&") and (title Like '%"& keyword &"%' or title like title Like '%"& keyword &"%' or title Like '%"& keyword &"%') order by articleid desc"
end if   
set rs=conn.execute(sql) 
if keyword="" then
   response.write"查找字符不能为空串，请重输入查找的信息<a href=""javascript:history.go(-1)"">返回重查</a>"
   response.end
 elseif rs.eof then
   response.write"没有你要查找的信息<a href=""javascript:history.go(-1)"">返回重查</a>"
   response.end
else   
  Set rs= Server.CreateObject("ADODB.Recordset")
  rs.open sql,conn,1,1
end if  
%>
<%
   const MaxPerPage=10
   dim totalPut   
   dim CurrentPage
   dim TotalPages
   dim i,j
   dim typename
      typename="搜索结果"
   if not isempty(request("page")) then
      currentPage=cint(request("page"))
   else
      currentPage=1
   end if
       set rst=server.CreateObject("ADODB.RecordSet") 
   %>
<body>
<script type="text/javascript">
 var Sys = {};
 var ua = navigator.userAgent.toLowerCase();
 var s;
 (s = ua.match(/msie ([\d.]+)/)) ? Sys.ie = s[1] :
 (s = ua.match(/firefox\/([\d.]+)/)) ? Sys.firefox = s[1] :
 (s = ua.match(/chrome\/([\d.]+)/)) ? Sys.chrome = s[1] :
 (s = ua.match(/opera.([\d.]+)/)) ? Sys.opera = s[1] :
 (s = ua.match(/version\/([\d.]+).*safari/)) ? Sys.safari = s[1] : 0;

//以下进行测试
 if (Sys.ie) document.write('<div style="top:2px;left:0;position:absolute; width:100%; height:100%;">');
 if (Sys.firefox) document.write('<div style="top:2px;position:absolute; width:100%; height:100%;z-index:-1;">');
 if (Sys.chrome) document.write('<div style="top:2px;position:absolute; width:100%; height:100%;z-index:-1;">');
 if (Sys.opera) document.write('<div style="top:2px;position:absolute; width:100%; height:100%;z-index:-1;">');
 if (Sys.safari) document.write('<div style="top:2px;position:absolute; width:100%; height:100%;z-index:-1;">');
</script>
<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0" width="100%" height="100%">
              <param name="movie" value="images/spider.swf" />
              <param name="quality" value="low" />
              <param name="wmode" value="transparent" />
              <param name="SCALE" value="exactfit">
              <embed src="images/spider.swf" width="100%" type="application/x-shockwave-flash" height="100%" quality="low" pluginspage="http://www.macromedia.com/go/getflashplayer" wmode="transparent" scale="exactfit"></embed>
     
</object>
</div>
<div class="cont">
<!--#include file="top.asp"-->
    <div class="main">
        <table width="100%" border="0" cellpadding="10" cellspacing="0">
        <tr> 
        <td align="center" valign="top">
              <div>查找结果&gt;&gt; <%response.write ""&typename1&""%> </div>
          <%       
  if rs.eof and rs.bof then 
       response.write "<p align='center'> 还 没 有 任 何 图 片</p>" 
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
           showpage totalput,MaxPerPage,"search.asp" 
            showContent 
            showpage1 totalput,MaxPerPage,"search.asp" 
       else 
          if (currentPage-1)*MaxPerPage<totalPut then 
            rs.move  (currentPage-1)*MaxPerPage 
            dim bookmark 
            bookmark=rs.bookmark 
           showpage totalput,MaxPerPage,"search.asp" 
            showContent 
             showpage1 totalput,MaxPerPage,"search.asp" 
        else 
	        currentPage=1 
           showpage totalput,MaxPerPage,"search.asp" 
           showContent 
           showpage1 totalput,MaxPerPage,"search.asp" 
	      end if 
	   end if 
   rs.close 
   end if 
	         
   set rs=nothing   
   conn.close 
   set conn=nothing 
   sub showContent 
       dim i 
	   i=0 
        %> <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bordercolorlight="#000000" bordercolordark="#FFFFFF">
            <tr align="center" bgcolor="#E3F2FD"> 
              <th width="50%"  height="20">图 集 标 题</th>
              <th width="20%"> 点 击</th>
              <th width="30%"> 加 入 日 期</th>
            </tr>
            <%i=0 
          do while not rs.eof%>
            <tr height="20" bgcolor="#EFF8FE"> 
              <td>·<a href="photoDetail.asp?id=<%=rs("articleid")%>" target="_blank"><%=rs("title")%></a></td>
              <td align="center"><%=rs("hits")%> </td>
              <td align="center"><%=rs("dateandtime")%></td>
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
  Response.Write "<table border=""0"" cellspacing=""2"" width=""100%"" bgcolor=""#F1F1F1"">"
      response.write  "<form action=""search.asp"" method=""post"" name=""search"" onsubmit=""return checkinput()"">"  
  response.write "<tr><td>" 
  Response.Write "共有图片：<font face=""arial"" color=""red"">"&totalnumber&"</font>篇 <font face=""arial"" color=""red"">"&CurrentPage&"/"&n&"</font>页" 

  if CurrentPage<2 then 
    response.write "<font color='999966'>首页 上一页</font>&nbsp;" 
  else 
    response.write "<a href="&filename&"?typeid="&typeid&"&page=1&>首页</a>&nbsp;" 
    response.write "<a href="&filename&"?typeid="&typeid&"&page="&(CurrentPage-1)&"&keyword="&keyword
	response.write ">上一页</a>&nbsp;" 
  end if 
  if n-currentpage<1 then 
    response.write "<font color='999966'>下一页 尾页</font>" 
  else 
       response.write "<a href="&filename&"?typeid="&typeid&"&page="&(CurrentPage+1)&"&keyword="&keyword
    response.write ">下一页</a> <a href="&filename&"?typeid="&typeid&"&page="&n&">尾页</a>" 
  end if 
 
 response.write   "<input type=""text"" name=keyword size=14 maxlength=""30"" value="""&keyword&""" onFocus=this.value=''>" 
  Response.Write  "<input type=""hidden"" name=""typeid"" value="&typeid&">" 
  response.write	" <input type=""Submit"" name=""search"" value=""搜 索"" title=""查询"">" 

  Response.Write "</td></tr>"
  response.write "</form>"
  Response.Write"</table>" 

end function  
function showpage1(totalnumber,maxperpage,filename) 
  dim n 
  if totalnumber mod maxperpage=0 then 
     n= totalnumber \ maxperpage 
  else 
     n= totalnumber \ maxperpage+1 
  end if 
  Response.Write "<table border=""0"" cellspacing=""2"" width=""100%"">" 
  response.write "<tr><td align=""left""><td width=""64%"">" 
  gotopage 
  response.write "<td width=""36%""><div align=""right"">" 
  if CurrentPage<2 then 
    response.write "<font color='999966'>首页 上一页</font>&nbsp;" 
  else 
    response.write "<a href="&filename&"?typeid="&typeid&"&page=1&>首页</a>&nbsp;" 
    response.write "<a href="&filename&"?typeid="&typeid&"&page="&(CurrentPage-1)&"&keyword="&keyword
	response.write ">上一页</a>&nbsp;" 
  end if 
  if n-currentpage<1 then 
    response.write "<font color='999966'>下一页 尾页</font>" 
  else 
       response.write "<a href="&filename&"?typeid="&typeid&"&page="&(CurrentPage+1)&"&keyword="&keyword
    response.write ">下一页</a> <a href="&filename&"?typeid="&typeid&"&page="&n&">尾页</a>" 
  end if 
  Response.Write "</div></td></tr></table>" 
end function  
sub gotopage%> <form action="<%response.write ""&filename&"?typeid="&typeid&"&keyword="&keyword&""%>" method="POST" id="form2" name="form2">
            <p>转到
              <input type="Text" name="page" value="1" size="4" maxlength="4">
              页 
              <input type="Submit" name="go" value="GO" title="执行">
            </p>
          </form>
          <% 
end sub 
        %></td>
      </tr>
    </table>
    </div>
</div>
<!--#include file="foot.asp"-->