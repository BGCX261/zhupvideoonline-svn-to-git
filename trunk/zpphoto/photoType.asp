<!--#include file="dbConn.asp"-->
<!--#include file="head.asp"-->
<%
   dim rs
   dim rst
   MaxPerPage=12
   dim totalPut   
   dim CurrentPage
   dim TotalPages
   dim i,j,b

   dim ty
   typeid=request("typeid")
   if not isempty(request("page")) then
      currentPage=cint(request("page"))
   else
      currentPage=1
   end if
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
 else typename1="全部图片"	
 end if
rstype.close
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
    <div class="left">
      <ul class="paihang">
          <li class="paihangtitle"> :: 本类浏览排行榜</li>
                  <%if typeid=0 then                      
    sql="select articleid,title,dateandtime,hits from albums order by hits desc"                      
else                      
   sql="select articleid,title,dateandtime,hits from albums where typeid="&typeid&" order by hits desc"                      
end if                         
Set rst= Server.CreateObject("ADODB.Recordset")                      
rst.open sql,conn,1,1
j=0
do while not rst.eof
j=j+1%>
                <li><a title="加入时间：<%=rst("dateandtime")%>" href="photoDetail.asp?id=<%=rst("articleid")%>"><%=rst("title")%></a></li>  <%
if j=20 then exit do
rst.movenext
loop
rst.close
set rst=hothing 
%> 

            </ul>
			</div>
    <div class="right">  <%                      
dim sql                               
if typeid=0 then                      
    sql="select articleid,title,bestpic,dateandtime,author,lx,hits from albums order by dateandtime desc"                      
else                      
    sql="select articleid,title,bestpic,dateandtime,author,lx,hits from albums where typeid=" & typeid & " order by dateandtime desc"                      
end if                         
Set rs= Server.CreateObject("ADODB.Recordset")                      
rs.open sql,conn,1,1                      
 if rs.eof and rs.bof then                      
       response.write "<p align='center'> 还 没 有 任 何 图 集</p>"                      
   else                      
	 ' totalPut=rs.recordcount                      
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
           showpage totalput,MaxPerPage,"photoType.asp"                      
            showContent                      
            showpage1 totalput,MaxPerPage,"photoType.asp"                      
       else                      
          if (currentPage-1)*MaxPerPage<totalPut then                      
            rs.move  (currentPage-1)*MaxPerPage                      
            dim bookmark                      
            bookmark=rs.bookmark                      
           showpage totalput,MaxPerPage,"photoType.asp"                      
            showContent                      
            showpage1 totalput,MaxPerPage,"photoType.asp"                      
        else                      
	        currentPage=1                      
          showpage totalput,MaxPerPage,"photoType.asp"                      
           showContent                      
           showpage1 totalput,MaxPerPage,"photoType.asp"                      
	      end if                      
	   end if                      
   end if                      
                       
                        
                      
   sub showContent                      
       dim i                      
	   i=0  %>                    
            <div class="pClass">            
   <% do while not rs.eof
    i=i+1
            %>
               <dl class="smlpicbox">
          <dt class="smlpic">
<%
  if rs("lx")=0 then
  %>
                        <a href='photoDetail.asp?id=<%= rs("articleid") %>'><img src='bestpics.asp?id=<%= rs("articleid") %>' onerror=""this.src='images/nopic.gif'"" /></a> 
                        <%
  elseif rs("lx")=1 then
   %>
   <a href='photoDetail.asp?id=<%= rs("articleid") %>'><img src='bestpics.asp?id=<%= rs("articleid") %>' onerror=""this.src='IMAGES/nopic.gif'"" /></a>
   <%
  end if
%>
</dt>

                      <dd class="pictitle"> <%=rs("title")%>  点击：<%=rs("hits")%></dd>
                    </dl>


                <%

    if i>=12 then exit do
    rs.movenext
    loop
		    %>
			</div>
                <form action="<%response.write ""&filename&"?typeid="&typeid&""%>" method="get">
                  <div class="pinfo"> 分页： 
                    <%            
   end sub             
function showpage(totalnumber,maxperpage,filename)            
  dim n            
  if totalnumber mod maxperpage=0 then            
     n= totalnumber \ maxperpage            
  else            
     n= totalnumber \ maxperpage+1            
  end if                     
  response.write "<div class=""pinfo"">"            
  Response.Write "本类共有图集：<font color=""red"">"&totalnumber&"</font>篇 <font color=""red"">"&CurrentPage&"/"&n&"</font>页"                                   	            
  response.write "&nbsp;&nbsp;"            
  if CurrentPage<2 then            
    response.write "首页 上一页&nbsp;"            
  else            
    response.write "<a href="&filename&"?typeid="&typeid&"&page=1&>首页</a>&nbsp;"            
    response.write "<a href="&filename&"?typeid="&typeid&"&page="&CurrentPage-1&">上一页</a>&nbsp;"            
  end if            
  if n-currentpage<1 then            
    response.write "下一页 尾页"            
  else            
       response.write "<a href="&filename&"?typeid="&typeid&"&page="&(CurrentPage+1)            
    response.write ">下一页</a> <a href="&filename&"?typeid="&typeid&"&page="&n&">尾页</a>"            
  end if            
  Response.Write "</div>"                      
end function             
function showpage1(totalnumber,maxperpage,filename)
  dim n
  if totalnumber mod maxperpage=0 then
     n= totalnumber \ maxperpage
  else
     n= totalnumber \ maxperpage+1
  end if

if rs.recordcount<= 12 then   
for j=1 to 1
response.write "[<a href="&filename&"?typeid="&typeid&"&page="&j&">"&j&"</a>]"   
next   
else
if n>10 then
b=n
n=10
end if   
for j=1 to n
response.write "[<a href="&filename&"?typeid="&typeid&"&page="&j&">"&j&"</a>]"   
next   
if b>10 then
response.write  "[<a href="&filename&"?typeid="&typeid&"&page="&CurrentPage+1&"' title='下一页'>&gt;&gt;</a>]"   
end if
end if 
end function             
sub gotopage
end sub         
          
           
    %> 
                    <input type="Text" name="page" size="4" maxlength="4"> 
                    <input type="submit" value="GO"> 
                
            </div>
			</form>
   <%
   rs.close                                                    
   set rs=nothing                        
   %> 
        </div>
    </div>
</div>
<!--#include file="foot.asp" -->