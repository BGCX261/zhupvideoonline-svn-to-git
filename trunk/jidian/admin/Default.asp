<!--#include file="bsconfig.asp"-->

<html><head>
<meta http-equiv=Content-Type content=text/html;charset=gb2312>
<title>∷网站后台管理∷</title>
<style type="text/css">
.navPoint {COLOR:#000; CURSOR: hand; FONT-FAMILY: Webdings; FONT-SIZE:12px;}
.a2{BACKGROUND: A4B6D7;}
</style>
</head>
<%
select case Request("menu")
case ""
index
case "top"
top2

end select
%>
<% sub top2 %>
<BODY leftMargin=0 topMargin=0 rightMargin=0>
<CENTER>
<TABLE style="BORDER-COLLAPSE: collapse" cellSpacing=0 cellPadding=0 width="100%" border=0>
<TR>
<TD align=center width="100%" height=25 style="BACKGROUND:#039; COLOR: #fff; font-size:16px; font-weight:bold;"
><B>网站后台管理</B></TD>
</TR>
</TABLE>
</CENTER>
</body>
<% end sub %>

<% sub index %>
<body style="MARGIN: 0px" scroll=no onResize=javascript:parent.carnoc.location.reload()>
<script>
if(self!=top){top.location=self.location;}
function switchSysBar(){
if (switchPoint.innerText==3){
switchPoint.innerText=4
document.getElementById("frmTitle").style.display="none"
}else{
switchPoint.innerText=3
document.getElementById("frmTitle").style.display=""
}}
</script>

<table border="0" cellPadding="0" cellSpacing="0" height="100%" width="100%">
  <tr>
      <td colspan="3" height="24">		<iframe frameborder="0" id="main" name="top" scrolling="no" src="?menu=top" style="HEIGHT: 24; VISIBILITY: inherit; WIDTH: 100%; Z-INDEX: 1">
		</iframe></td>
  </tr>
  <tr>
    <td align="middle" noWrap vAlign="center" id="frmTitle">
    <iframe frameBorder="0" id="carnoc" name="carnoc" scrolling=no src="Menu_left.asp" style="HEIGHT: 100%; VISIBILITY: inherit; WIDTH: 170px; Z-INDEX: 2">
    </iframe>
    </td>
    <td style="WIDTH: 12px; border-left:#A4B6D7 1px solid;">
    <table border="0" cellPadding="0" cellSpacing="0" height="100%">
      <tr>
        <td style="HEIGHT: 100%;FONT-SIZE:12px; CURSOR: default; COLOR: #000;" onClick="switchSysBar()">
        <span class="navPoint" id="switchPoint" title="关闭/打开左栏">3</span><br>
        屏幕切换 </td>
      </tr>
    </table>
    </td>
		<td style="WIDTH: 100%;">
		<iframe frameborder="0" id="main" name="main" scrolling="yes" src="sysadmin_view.asp" style="HEIGHT: 96%; VISIBILITY: inherit; WIDTH: 100%; Z-INDEX: 1">
		</iframe></td>
  </tr>
</table>
<script>
if(window.screen.width<'1024'){switchSysBar()}
</script>
</body>
<%
end sub

Function IsObjInstalled(strClassString)
On Error Resume Next
IsObjInstalled = False
Err = 0
Dim xTestObj
Set xTestObj = Server.CreateObject(strClassString)
If 0 = Err Then IsObjInstalled = True
Set xTestObj = Nothing
Err = 0
End Function

%></html>