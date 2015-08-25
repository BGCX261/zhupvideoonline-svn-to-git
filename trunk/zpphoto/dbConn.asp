<% 
dim sql_leach,sql_leach_0,Sql_DATA,SQL_Get,Sql_Post 
sql_leach = "',and,exec,insert,select,delete,update,count,chr,mid,master,truncate,char,declare,0x4445434C4152452040542056415243484,4152452040542056415243484152283235352,0x4445434C415245204054205,48415228323535292C4043205641524,22832353529204445434C415245205,61626C655F437572736F7220435552534F5220464F52205345,45435420612E6E616D652C622E6E616D652046524F4D207379736F626A656, <script,92C404320564152434841522832353529204445434C415245205461626C655F437572736F7220435,F626A6563747320612C737973636F6C756D6E73206220574845524520612E69643D622E696420414E4420612E78747970653D27752720414E442028622E78747970653D3939204F5220622E78747970,5204054205641524348415228323535292C404320564152434,VaRcHaR,sYsDaTaBaSeS,select,dEcLaRe" 
'以上是除了常规防注入字符集外。又加了本木马专有的十六进制字符集。 
sql_leach_0 = split(sql_leach,",") 

If Request.QueryString <>"" Then 
For Each SQL_Get In Request.QueryString 
For SQL_Data=0 To Ubound(sql_leach_0) 
if instr(Request.QueryString(SQL_Get),sql_leach_0(Sql_DATA))>0 Then 
Response.Write "请不要提交可注入的字符串！非法字符为"&sql_leach_0(Sql_DATA) 
Response.end 
end if 
next 

Next 
End If 


If Request.Form <>"" Then 
For Each Sql_Post In Request.Form 
For SQL_Data=0 To Ubound(sql_leach_0) 
if instr(Request.Form(Sql_Post),sql_leach_0(Sql_DATA))>0 Then 
Response.Write "请不要提交可注入的字符串！非法字符为"&sql_leach_0(Sql_DATA) 
Response.end 
end if 
next 
next 
end if
%>
<%
   dim conn   
   dim connstr
   on error resume next
   connstr="DBQ="+server.mappath("z97122p.asp")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
     set conn=server.createobject("ADODB.CONNECTION")
     conn.open connstr  
%>


