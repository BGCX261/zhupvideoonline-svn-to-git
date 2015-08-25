<!--#include file="Inc/syscode.asp"-->
<!-- #include file="Inc/Head.asp" -->
  <div id="body">
    <table class="bodytable">
      <tr>
        <td class="bodyleft">
		  <!-- #include file="L_stuZone.asp" -->
        </td>
        <td class="bodyline"></td>
        <td class="bodyright">
        <!-- #include file="Inc/headAd.asp" -->
          <div class="mbox">
              <div class="artlist">
<%
MaxPerPage=10 
			
nbtyped=request("nbtyped")
if not isempty(request("page")) then
      currentPage=cint(request("page"))
else
      currentPage=1
end if
if not isEmpty(request("nbtyped")) then
  nbtyped=request("nbtyped")
else
  nbtyped=""
end if
if nbtyped="" then                      
    sql="select * from StuZone where Passed=True order by id desc"                      
else                      
    sql="select * from StuZone where nbtyped='"&request("nbtyped")&"' and Passed=True order by id desc"                    
end if 
   
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1

			         
	response.write  "<div class='artpath'><img src='img/title_ico.gif' /> <a href='stuZone.asp'>师生天地</a> >> "
	if nbtyped="" then
		response.write "所有信息"
	else
		if nbtyped<>"" then
			response.write "<a href='stuZone.asp?nbtyped=" & nbtyped & "'>" & nbtyped & "</a> "
		end if
	end if
    response.write  "</div>"

if rs.eof and rs.bof then                      
       response.write "<div align='center'>暂无数据</div>"                      
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
           showContent                      
           showpage totalput,MaxPerPage,"stuZone.asp"           
       else                      
          if (currentPage-1)*MaxPerPage<totalPut then                      
            rs.move  (currentPage-1)*MaxPerPage                      
            dim bookmark                      
            bookmark=rs.bookmark                                     
            showContent                      
            showpage totalput,MaxPerPage,"stuZone.asp"
        else                      
	        currentPage=1                                 
           showContent                      
           showpage totalput,MaxPerPage,"stuZone.asp" 
	      end if                      
	   end if                      
   end if                    

  
%>
<% sub showContent                      
   dim i                      
   i=0 %>
                <ul>
				  <%do while not rs.eof
                  i=i+1%>
				  <li>
                     <span class="artlnk"><a href="stuZoneInfo.asp?id=<%=rs("id")%>"><%=cutstr(rs("Title"),50)%></a></span>
					 <span class="artlistupdate"><%=formatdatetime(rs("UpdateTime"),2)%></span>
                  </li>
<% 
    if i>=10 then exit do
    rs.movenext
    loop
%>             
                </ul>
			  </div>
              <div id="tpage"><%end sub%>
                <%
function showpage(totalnumber,maxperpage,filename)            
  dim n            
  if totalnumber mod maxperpage=0 then            
     n= totalnumber \ maxperpage            
  else            
     n= totalnumber \ maxperpage+1            
  end if
  Response.Write "<img src='img/arrow.gif' />"                                 
  Response.Write "本类共有：<span id='numnews'>"&totalnumber&"</span>个 <span id='numnews'>"&CurrentPage&"/"&n&"</span>页"                                   	                       
  if CurrentPage<2 then            
    response.write "<span id='gray'>首页 上一页 </span>"            
  else            
    response.write "<span id='light'><a href="&filename&"?nbtyped="&nbtyped&"&page=1&>首页</a></span>"            
    response.write "<span id='light'><a href="&filename&"?nbtyped="&nbtyped&"&page="&CurrentPage-1&">上一页</a></span>"            
  end if            
  if n-currentpage<1 then            
    response.write "<span id='gray'>下一页 尾页 </span>"            
  else            
       response.write "<span id='light'><a href="&filename&"?nbtyped="&nbtyped&"&page="&(CurrentPage+1)            
    response.write ">下一页</a></span> <span id='light'><a href="&filename&"?nbtyped="&nbtyped&"&page="&n&">尾页</a></span>"            
  end if                                 
end function
rs.close                                                    
   set rs=nothing                                            
%>
              </div>
          </div>
        </td>
      </tr>
    </table>
  </div>
  <!-- #include file="Inc/Foot.asp" -->
</body>
</html>
