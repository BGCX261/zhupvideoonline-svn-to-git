
<!--#include file="bsconfig.asp"-->
<!-- #include file="Inc/Head.asp" -->
<%call checkmanage("19")%>
<%
'ASP ��������,���Ƶĺ���.���б��ֶΰ���ID ,sortId.
'sortIdΪ���������ֶ�. ����Ϊ��ID���,���ʱ��IDֵ����sortId

'���Ʋ���ʵ��Ϊ,������¼����һ����¼��sortIdֵ����,��֮��Ȼ.

id = request("id")      '��ǰ��¼id  
act = trim(request("act"))  '�ƶ�����
tempo = trim(request("tempo"))  '�ƶ�����

if act<>"" then Call MoveOrders(act,currentPage)

Sub MoveOrders(act,currentPage)

if NOT IsEmpty(request("page")) Then
currentPage=Cint(request("page"))
end if

dim temporders_a,temporders_b
if act="up" then   
   set rs=server.CreateObject("adodb.recordset")
   sqltxt="select sortId from Teacher where id=" & id 
   rs.open sqltxt,conn,1,1
   temporders_a = rs("sortId")
   'response.write "<script>alert('"& temporders_a&"��');</ script>"
   set rs2=server.CreateObject("adodb.recordset")
   rs2.open "select top 1 sortId,id from Teacher where sortId>" & temporders_a & " order by sortId asc",conn,1,3
   if rs2.eof then 
    response.write "<script>alert('�Ѿ�����һλ�ˣ�');history.go(-1)</script>"
   
   else
    temporders_b= rs2("sortId")
    response.write "<script>alert('"& temporders_b&"��');</script>"
    temp_b_id = rs2("id")
    rs2.close
    set rs2=nothing
    conn.execute("update Teacher set sortId=" & temporders_b &" where id="& id)
    conn.execute("update Teacher set sortId=" & temporders_a &" where id="& temp_b_id)
	
	if tempo = 1 then
       currentPage=currentPage-1
    else
       currentPage=currentPage
    end if
	
    response.write "<script>alert('���Ƴɹ���');</script>"
   end if  
else
   set rs=server.CreateObject("adodb.recordset")
   rs.open "select sortId from Teacher where id=" & id,conn,1,1
   temporders_a = rs("sortId")
   rs.close
   rs.open "select top 1 sortId,id from Teacher where sortId<" & temporders_a & " order by sortId desc",conn,1,3
   if rs.eof then 
    response.write "<script>alert('�Ѿ������һλ�ˣ�');history.go(-1)</script>"
   
   else
    temporders_b= rs("sortId")
    temp_b_id = rs("id")
    rs.close
    set rs=nothing
    conn.execute("update Teacher set sortId=" & temporders_b &" where id="& id)
    conn.execute("update Teacher set sortId=" & temporders_a &" where id="& temp_b_id)
	
	if tempo mod 10 = 0 then
       currentPage=currentPage+1
    else
       currentPage=currentPage
    end if
	
    response.write "<script>alert('���Ƴɹ���');</script>"
   end if  
end if
response.Redirect("Bs_coursTeacher.asp?page="&currentPage)
End Sub
 
%>
