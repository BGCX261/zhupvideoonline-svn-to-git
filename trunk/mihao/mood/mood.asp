<%
'���������ݿ�����������ݿ�����վ���ۺϿ����� <--#include file="../inc/conn.asp"-->��������������ݿ�����
dim conn,connstr
on error resume next
connstr="DBQ="+server.mappath("#mo od.asa")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open connstr

action = Replace(Trim(Request.QueryString("action")),"'","")
id = Replace(Trim(Request.QueryString("id")),"'","")
classid = Replace(Trim(Request.QueryString("classid")),"'","")
typee = Replace(Trim(Request.QueryString("typee")),"'","")


if action="show" then
	  set rs=server.createobject("adodb.recordset")
	  sql="Select * From Mood Where classid='"&classid&"' and cstr(moodid)='"&id&"'"
      rs.open sql,conn,1,1
	  If Not(rs.Eof And rs.Bof) Then
      mood1=rs("mood1")
      mood2=rs("mood2")
      mood3=rs("mood3")
      mood4=rs("mood4")
      mood5=rs("mood5")
      mood6=rs("mood6")
      mood7=rs("mood7")
      mood8=rs("mood8")
	  Rs.Close
	  Set Rs = Nothing
      Response.Write(""&mood1&","&mood2&","&mood3&","&mood4&","&mood5&","&mood6&","&mood7&","&mood8&"")
      else
	  conn.execute("Insert into Mood(classid,moodid)values('"&classid&"','"&id&"')")
      Response.Write("0,0,0,0,0,0,0,0")
	  end if
else
if action="mood" then
Select case typee
case "mood1"
conn.execute("Update Mood set mood1=mood1+1 where classid='"&classid&"' and cstr(moodid)='"&id&"'")
case "mood2"
conn.execute("Update Mood set mood2=mood2+1 where classid='"&classid&"' and cstr(moodid)='"&id&"'")
case "mood3"
conn.execute("Update Mood set mood3=mood3+1 where classid='"&classid&"' and cstr(moodid)='"&id&"'")
case "mood4"
conn.execute("Update Mood set mood4=mood4+1 where classid='"&classid&"' and cstr(moodid)='"&id&"'")
case "mood5"
conn.execute("Update Mood set mood5=mood5+1 where classid='"&classid&"' and cstr(moodid)='"&id&"'")
case "mood6"
conn.execute("Update Mood set mood6=mood6+1 where classid='"&classid&"' and cstr(moodid)='"&id&"'")
case "mood7"
conn.execute("Update Mood set mood7=mood7+1 where classid='"&classid&"' and cstr(moodid)='"&id&"'")
case else
conn.execute("Update Mood set mood9=mood8+1 where classid='"&classid&"' and cstr(moodid)='"&id&"'")
End Select
	  set rs=server.createobject("adodb.recordset")
	  sql="Select * From Mood Where classid='"&classid&"' and cstr(moodid)='"&id&"'"
      rs.open sql,conn,1,1
      mood1=rs("mood1")
      mood2=rs("mood2")
      mood3=rs("mood3")
      mood4=rs("mood4")
      mood5=rs("mood5")
      mood6=rs("mood6")
      mood7=rs("mood7")
      mood8=rs("mood8")
	  Rs.Close
	  Set Rs = Nothing
      Response.Write(""&mood1&","&mood2&","&mood3&","&mood4&","&mood5&","&mood6&","&mood7&","&mood8&"")
else
      Response.Write("0,0,0,0,0,0,0,0")
end if
end if
 %>