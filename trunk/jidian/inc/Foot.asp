<div id="footer">
  <div class="footbox">
	  <!--地址：<%=BsAddress%> | 邮编：<%=BsZipcode%> | Email：<a href="mailto:<%=BsEmail%>"><%=BsEmail%></a>电话：<%=BsPhone%> | 传真：<%=BsFax%>  <br> 版权所有：<%=BsCompanyName%> <br>--> 
	<%
    copyRight=replace(rs_BsCo("copyRight"),"../","./")
    response.write copyRight
    %>
  </div>
</div>
<%
rs_BsCo.close
rs.close
set rs=nothing
conn.close
set conn=nothing  
%>