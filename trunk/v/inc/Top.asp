<%
'��վ����������
sql="select * from main"
Set rs_main= Server.CreateObject("ADODB.Recordset")
rs_main.open sql,conn,1,1
Profile=ubbcode(rs_main("Profile")) '��˾���
ADSheet=ubbcode(rs_main("ADSheet")) '�������
Demoflv=ubbcode(rs_main("Demoflv")) 'ʹ����ʾ����
Declare=ubbcode(rs_main("Declare")) 'ҳ������
Review=rs_main("Review") '�����Ƿ���Ҫ���
rs_main.close
%>