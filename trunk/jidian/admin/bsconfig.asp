<!--#include file="Conn.asp"-->
<%
dim Name
Name=replace(session("Name"),"'","")
if Name="" then
	call CloseConn()
response.Write("<script language='javascript'>top.location='Login.asp';</script>")
response.end
end if
%>

<%
sub htmlend
%>
    <div align="center">Script Execution Time:<%=fix((timer()-startime)*1000)%>ms </div>
</body>
</html>
<%
responseend
end sub
%>
