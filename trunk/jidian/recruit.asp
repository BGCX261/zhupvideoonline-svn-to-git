<!--#include file="Inc/syscode.asp"-->
<!-- #include file="Inc/Head.asp" -->
  <div id="body">
    <table class="bodytable">
      <tr>
        <td class="bodyleft">
            <!-- #include file="L_recruit.asp" -->
        </td>
        <td class="bodyline"></td>
        <td class="bodyright">
        <!-- #include file="Inc/headAd.asp" -->
          <div class="mbox">
            <div class="artlist">
<%
MaxPerPage=10 
			
typed=request("typed")
if not isempty(request("page")) then
      currentPage=cint(request("page"))
else
      currentPage=1
end if
if not isEmpty(request("typed")) then
  typed=request("typed")
else
  typed=""
end if
if typed="" then                      
    sql="select * from Recruit order by id desc,mdy_time desc"                      
else                      
    sql="select * from Recruit where typed='"&request("typed")&"' order by id desc,mdy_time desc"                    
end if 
   
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1

			         
	response.write  "<div class='artpath'><img src='img/title_ico.gif' /> <a href='recruit.asp'>招生就业</a> >> "
	if typed="" then
		response.write "所有信息"
	else
		if typed<>"" then
			response.write "<a href='recruit.asp?typed=" & typed & "'>" & typed & "</a> "
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
           showpage totalput,MaxPerPage,"recruit.asp"           
       else                      
          if (currentPage-1)*MaxPerPage<totalPut then                      
            rs.move  (currentPage-1)*MaxPerPage                      
            dim bookmark                      
            bookmark=rs.bookmark                                     
            showContent                      
            showpage totalput,MaxPerPage,"recruit.asp"
        else                      
	        currentPage=1                                 
           showContent                      
           showpage totalput,MaxPerPage,"recruit.asp" 
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
				  <li  <%if rs("toTop")=true then response.write "class='toTop'" else response.write "class=''" end if%>>
                    <span><a href="recruitInfo.asp?id=<%=rs("id")%>"><%=cutstr(rs("Title"),50)%></a></span>
                    <span class="artlistupdate"><%=rs("time")%></span>
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
    response.write "<span id='light'><a href="&filename&"?typed="&typed&"&page=1&>首页</a></span>"            
    response.write "<span id='light'><a href="&filename&"?typed="&typed&"&page="&CurrentPage-1&">上一页</a></span>"            
  end if            
  if n-currentpage<1 then            
    response.write "<span id='gray'>下一页 尾页 </span>"            
  else            
       response.write "<span id='light'><a href="&filename&"?typed="&typed&"&page="&(CurrentPage+1)            
    response.write ">下一页</a></span> <span id='light'><a href="&filename&"?typed="&typed&"&page="&n&">尾页</a></span>"            
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
