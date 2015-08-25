<%
'打开站点设置连接
sql="select * from main"
Set rs_main= Server.CreateObject("ADODB.Recordset")
rs_main.open sql,conn,1,1
notice=ubbcode(rs_main("notice"))
noticeTime=rs_main("noticeTime")
sent=ubbcode(rs_main("sent"))
rs_main.close
%>