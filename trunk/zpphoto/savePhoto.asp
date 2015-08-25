<!--#include file="security.asp"-->
<!--#include file="inc/articlechar.inc"-->
<!--#include file="dbConn.asp"-->
<html>
<%
Set rs = Server.CreateObject("ADODB.Recordset")
%>
<%
if request.form("title")="" or request.form("content")="" or request.form("minPic")="" or request.form("typeid")=0 then 
response.write"<script>alert('表单未填完整，请返回重填！');history.back();</Script>"
response.end
end if
%>
<%
if request.form("html")="on" then
content=request.form("content")
else
content=htmlencode2(request.form("content"))
end if


sql="select * from albums" 

rs.open sql,conn,3,2
rs.addnew
rs("typeid")=Request.Form("typeid")
rs("orgTime")=request.form("QYear")&"年"&request.form("QMonth")&"月"
rs("images")=request.form("images")
rs("images1")=request.form("images1")
rs("images2")=request.form("images2")
rs("images3")=request.form("images3")
rs("images4")=request.form("images4")
rs("images5")=request.form("images5")
rs("images6")=request.form("images6")
rs("images7")=request.form("images7")
rs("images8")=request.form("images8")
rs("images9")=request.form("images9")
rs("images10")=request.form("images10")
rs("images11")=request.form("images11")
rs("images12")=request.form("images12")
rs("images13")=request.form("images13")
rs("images14")=request.form("images14")
rs("images15")=request.form("images15")
rs("images16")=request.form("images16")
rs("images17")=request.form("images17")
rs("images18")=request.form("images18")
rs("images19")=request.form("images19")
rs("images20")=request.form("images20")
rs("bestpic")=request.form("minPic")
rs("best")=request.form("best")
rs("title")=request.form("title")
rs("content")=content
rs("dateandtime")=now()
rs("author")=request.form("author")
rs("source")=request.form("source")
rs("lx")=request.form("lx")
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<head>
<title>图集保存</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" type="text/css" href="css/style.css">
</head>

<body>
<div class="admCont">
  <table class="iptTable">
    <tr> 
      <td class="doTitle">添加图片成功</td>
    </tr>
    <tr> 
      <td align="center"> <br>
          标题为：<%=request.form("title")%>
          是否继续添加？<br>
          <br>
          <a href="addPhoto.asp">是</a>&nbsp;&nbsp; <a href="managePhoto.asp">否</a><br>
          <br>
      </td>
    </tr>
  </table>
</div>
</body>
</html>


