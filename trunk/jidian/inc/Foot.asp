<div id="footer">
  <div class="footbox">
	  <!--��ַ��<%=BsAddress%> | �ʱࣺ<%=BsZipcode%> | Email��<a href="mailto:<%=BsEmail%>"><%=BsEmail%></a>�绰��<%=BsPhone%> | ���棺<%=BsFax%>  <br> ��Ȩ���У�<%=BsCompanyName%> <br>--> 
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