<div class="footDeclare"><%=Declare%></div>
<div id="footer">
  <div class="footCopy">
    地址：<%=BsAddress%> | 邮编：<%=BsZipcode%> | 电话：<%=BsPhone%> | Email：<a href="mailto:<%=BsEmail%>"><%=BsEmail%></a><br>
    版权所有：<%=BsCompanyName%> | 备案序号：<a href="http://www.miibeian.gov.cn/" target="_blank"><%=BsICPcode%></a> <br />
  </div>
</div>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing  
%>
