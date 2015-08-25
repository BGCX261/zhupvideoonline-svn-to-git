<%
'打开站点设置连接
sql="select * from main"
Set rs_main= Server.CreateObject("ADODB.Recordset")
rs_main.open sql,conn,1,1
Profile=ubbcode(rs_main("Profile")) '公司简介
ADSheet=ubbcode(rs_main("ADSheet")) '广告宣传
Demoflv=ubbcode(rs_main("Demoflv")) '使用演示视屏
Declare=ubbcode(rs_main("Declare")) '页底申明
Review=rs_main("Review") '评论是否需要审核
rs_main.close
%>