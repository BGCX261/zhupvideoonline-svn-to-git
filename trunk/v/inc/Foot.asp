<div class="footDeclare"><%=Declare%></div>
<div id="footer">
  <div class="footCopy">
    ��ַ��<%=BsAddress%> | �ʱࣺ<%=BsZipcode%> | �绰��<%=BsPhone%> | Email��<a href="mailto:<%=BsEmail%>"><%=BsEmail%></a><br>
    ��Ȩ���У�<%=BsCompanyName%> | ������ţ�<a href="http://www.miibeian.gov.cn/" target="_blank"><%=BsICPcode%></a> <br />
  </div>
</div>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing  
%>
