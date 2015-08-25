<%
'生成公告.txt文件
Sub NoticeSite()
dim tmpStr
Sql="Select * From Notice WHERE EnableNotice=True order by OrderID DESC"
Openrs(sql)
For i = 1 to Rs.Recordcount
tmpStr=tmpStr&"marqueeContent["&i-1&"]='<a href="""&rs("Nurl")&""" target=""_blank"">"&rs("NTitle")&"</a>';"&vbcrlf&""
Rs.Movenext
Next
Rs.Close
Set Rs=nothing
Call Save("../include/Notice.txt",tmpStr)
End Sub
'========================================================================
'生成搜索热词
%>
