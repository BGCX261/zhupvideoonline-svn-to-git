<!--#include file="Inc/syscode.asp"-->
<!-- #include file="Inc/Head.asp" -->
<div id="body">
  <table class="bodytable">
    <tr>
      <td class="bodyleft">
	   <div id="leftbg">
          <!-- #include file="L_willbook.asp" -->
        </div></td>
      <td class="bodyline"></td>
      <td class="bodyright">
	  <table class="mbox">
          <tr>
            <td><div class="rtitle"><img src="img/title_ico.gif" /> �� Ƶ �� �� �� ��</div></td>
          </tr>
          <tr>
            <td class="cttd">
			<%
	MaxPerPage=12 
			
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
    sql="select * from newbook order by id desc"                      
else                      
    sql="select * from newbook where nbtyped='"&request("nbtyped")&"' order by id desc"                    
end if 
   
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1

			         
	response.write  "<div><a href='willbook.asp'>��Ƶ�����̳�</a> >>"
	if nbtyped="" then
		response.write "���н̳�"
	else
		if nbtyped<>"" then
			response.write "<a href='willbook.asp?nbtyped=" & nbtyped & "'>" & nbtyped & "</a> "
		end if
	end if
    response.write  "</div>"

if rs.eof and rs.bof then                      
       response.write "<td align='center'>��������</td>"                      
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
           showpage totalput,MaxPerPage,"willbook.asp"           
       else                      
          if (currentPage-1)*MaxPerPage<totalPut then                      
            rs.move  (currentPage-1)*MaxPerPage                      
            dim bookmark                      
            bookmark=rs.bookmark                                     
            showContent                      
            showpage totalput,MaxPerPage,"willbook.asp"
        else                      
	        currentPage=1                                 
           showContent                      
           showpage totalput,MaxPerPage,"willbook.asp" 
	      end if                      
	   end if                      
   end if                    

  
%>


<% sub showContent                      
       dim i                      
	   i=0 %>
                <table  cellspacing="8">
                      <tr>
				<%do while not rs.eof
                  i=i+1%>
                  <td style="border:#CCC solid 1px; width:200px; text-align:center;">
						
				  <table>
                      <tr>
                        <td align="center"><a href="willbookInfo.asp?id=<%=rs("id")%>"><img src="<%=rs("DefaultPicUrl")%>" width="100" border="0" /></a></td>
                      </tr>
					  <tr>
                        <td style="color:#999; text-align:center;">[<%=rs("nbtyped")%>]</td>
                      </tr>
                      <tr>
                        <td align="center"><div id="chpname"><a href="willbookInfo.asp?id=<%=rs("id")%>"><%=rs("Title")%></a></div></td>
                      </tr>
					  <tr><td  style="color:#999; text-align:center;"><%=rs("time")%> &nbsp; 
                  <%if rs("time")=date() then%>
                  <strong><font color="#FF0000">New</font></strong>
                  <%end if%></td></tr>
                    </table></td>
						 <%
	if (i mod (Maxperpage/3)=0) and i>=(Maxperpage/3) then
             %>
                      </tr>
                      <tr>
                        <%  
'rs.MoveNext      
'i=i+1      
'Loop
'end sub   			      
%>
                        <% end if
    if i>=12 then exit do
    rs.movenext
    loop

%>
                      </tr>
                    </table>
					
              </td>
          </tr>  
          <tr>
		    <td id="tpage">
			<%end sub%>
                <%
function showpage(totalnumber,maxperpage,filename)            
  dim n            
  if totalnumber mod maxperpage=0 then            
     n= totalnumber \ maxperpage            
  else            
     n= totalnumber \ maxperpage+1            
  end if
  Response.Write "<img src='img/arrow.gif' />"                                 
  Response.Write "���๲�У�<span id='numnews'>"&totalnumber&"</span>�� <span id='numnews'>"&CurrentPage&"/"&n&"</span>ҳ"                                   	                       
  if CurrentPage<2 then            
    response.write "<span id='gray'>��ҳ ��һҳ </span>"            
  else            
    response.write "<span id='light'><a href="&filename&"?nbtyped="&nbtyped&"&page=1&>��ҳ</a></span>"            
    response.write "<span id='light'><a href="&filename&"?nbtyped="&nbtyped&"&page="&CurrentPage-1&">��һҳ</a></span>"            
  end if            
  if n-currentpage<1 then            
    response.write "<span id='gray'>��һҳ βҳ </span>"            
  else            
       response.write "<span id='light'><a href="&filename&"?nbtyped="&nbtyped&"&page="&(CurrentPage+1)            
    response.write ">��һҳ</a></span> <span id='light'><a href="&filename&"?nbtyped="&nbtyped&"&page="&n&">βҳ</a></span>"            
  end if                                 
end function
rs.close                                                    
   set rs=nothing                                            
%>
                         </td>
          </tr>
        </table></td>
    </tr>
  </table>
</div>
<!-- #include file="Inc/Foot.asp" -->
</body>
</html>
