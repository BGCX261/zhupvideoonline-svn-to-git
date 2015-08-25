<!--#include file="include/head.asp"-->
<!--#include file="include/ofunctions.asp"-->
<%If User_Group<>"admin" Then TurnTo "/"%>
<%
Dim Form_ID,Form_context,Title,Content,poster,tags,ClassID,SubClass,SetTop,Classic,TitleBold,TitleColor,Lock,Hide,Form_Save,j,getpicRight,Verify

Act=Request.QueryString("act")
Form_ID=Request.QueryString("id")
Form_Save=Request.Form("save")

'读取子栏目
tempForum=ForumID
If ForumID="" Then tempForum=1
tempClass=Application(SystemKey&"subclass"&tempForum)
If instr(tempClass,"|")>0 Then
subclassArray=split(tempClass,"|")
Else
subclassArray=split("分类","|")
End If

Title=Request.Form("Title")
Content=Request.Form("Content")
Poster=Request.Form("Poster")
Tags=Request.Form("Tags")
ClassID=Request.Form("ForumID")
SubClass=Request.Form("SubClass")
SetTop=Request.Form("SetTop")
Classic=Request.Form("Classic")
TitleBold=Request.Form("TitleBold")
TitleColor=Request.Form("TitleColor")
Lock=Request.Form("Lock")
Hide=Request.Form("Hide")
Verify=Request.Form("Verify")
getpicRight=Cnum(Request.Form("getpicRight"))

If act="" Then
act="new"
Call PostEdit
End If

If act="edit" Then
act="EditText"
Sql="select * from Content where id="&Form_ID&" "
OpenRs(Sql)
forumID=Rs("classID")
subClassID=Rs("subclass")
Title=Rs("Title")
form_context=Rs("Content")
Poster=Rs("Poster")
Tags=Rs("Tags")
SetTop=Rs("SetTop")
Classic=Rs("Classic")
TitleBold=Rs("TitleBold")
TitleColor=Rs("TitleColor")
Lock=Rs("Lock")
Hide=Rs("Hide")
Verify=Rs("Verify")
Call PostEdit
End If

Sub PostEdit
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="admin/images/css.css" rel="stylesheet" type="text/css">
</head>
<body>
<div class="DB">
    <div class="L1">当前位置：<a href="admin/admin_show.asp">管理面板</a> &gt; 文章编辑</div>
    <div class="D2">
<form name="form1" method="POST" action="apost.asp?act=<%=act%>">
<input type="hidden" name="artid" value="<%=Form_ID%>">
<input type="hidden" name="save" value="yes"><span id='fff'></span>
<p class="P1">文章标题 <input type="text" name="title" class="IT1" maxlength="60" value="<%=Title%>">
<select class="Sel" name="TitleColor">
<option value="">标题颜色</option>
<option value="#ee1b2e" style="background-color:#ee1b2e;color:#ee1b2e;">红色</option>
<option value="#2b65b7" style="background-color:#2b65b7;color:#2b65b7;">蓝色</option>
<option value="#3c9d40" style="background-color:#3c9d40;color:#3c9d40;">绿色</option>
<option value="#2897c5" style="background-color:#2897c5;color:#2897c5;">淡蓝</option>
<option value="#ef8819" style="background-color:#ef8819;color:#ef8819;">橙色</option>
</select> <input type="checkbox" <%if TitleBold=True then Response.write "checked value=""1"" " Else Response.write "value=""1""" End if%> name="TitleBold">标题加粗</p>
<p class="P1">作者来源 <input name="Poster" class="IT2" type="text" value="<%=Poster%>">&nbsp;&nbsp;&nbsp;&nbsp;标签 <input name="Tags" type="text" class="IT1" value="<%=Tags%>"> 多个关键字用“|”隔开</p>
<p class="P1">
<%
'设置栏目名称
Response.write "隶属栏目 <select name=""forumID"" id=""classname"" onchange=""javascript:setclass();"" class=""Sel"">"
Response.write "<option value="""">设置栏目</option>"
For j=1 To BoardCount
Response.write "<option"
If j=Cnum(forumID) Then
Response.write " selected"
End If
Response.write " value="&j&">"&Application(systemkey&j)&"</option>"
Next
Response.write "</select>"
'显示分类
Response.write "&nbsp;<select name=""subclass"" id=""subclass"" class=""Sel""><option value="""">选择分类</option>"
For j=1 to ubound(subclassArray)
Response.write("<option ")
If j=Cnum(subclassID) Then response.write " selected "
Response.write("value="&j&">"&subclassArray(j)&"</option>")
Next
Response.write "</select>"
'分类结束
%>
</p>

<div class="divtext">
<textarea id="content" name="content" rows="12" cols="80" style="width:960px;height:300px;"><%=form_context%></textarea>
</div>
<p class="P2">

<%If act="EditText" And Verify="-1" Then
Response.Write ("<input type=""checkbox"" value=""0"" name=""Verify"">通过审核&nbsp;&nbsp;&nbsp;&nbsp;")
Else
Response.Write ("<input type=""hidden"" name=""Verify"" value=""0"">")
End If
%> 
<input type="checkbox" value="1" <%if SetTop=True then response.write "checked" end if%> name="SetTop">置顶&nbsp;&nbsp;&nbsp;&nbsp;<input type="checkbox" value="1" <%if Classic=True then response.write "checked" end if%> name="Classic">推荐&nbsp;&nbsp;&nbsp;&nbsp;<input type="checkbox" value="1" <%if Hide=True then response.write "checked" end if%> name="Hide">热点新闻&nbsp;&nbsp;&nbsp;&nbsp;<input type="checkbox" value="1" name="getpicRight">保存远程图片</p>
<p class="P3"><input type="submit" class="btn" value="确定发表"></p>
<script language="javascript" >
//分类自动更新
function setclass()
{
document.all("fff").innerHTML="";
var boardCount=<%=BoardCount%>
var subclass='<%=Application(SystemKey&"subclassAll")%>';
var pro_Class=new Array();
var pro_SubClass=subclass.split(',');

pro_Class[0]='--分类--';
pro_Class[boardCount+1]='<%=Application(Systemkey&"subclass999")%>';

for(z=1;z<=boardCount;z++) {
pro_Class[z]=pro_SubClass[z-1];
}
var nowSelectIndex=document.all("classname").selectedIndex;
for(i=document.all("subclass").length-1;i>=0;i--){document.all("subclass").options.remove(i);}

var Array_Class=pro_Class[nowSelectIndex].split('|');
   if(Array_Class.length>1){
   for(j=0;j<Array_Class.length;j++){
   document.all("subclass").options.add(new Option(Array_Class[j],j));
   document.all("subclass")[0].selected=true;
   }
  }else{
  document.all("subclass").options.add(new Option(pro_Class[nowSelectIndex],0));
  }
}
</script>
<script type="text/javascript" src="xheditor/jquery.js"></script>
<script type="text/javascript" src="xheditor/xheditor.js"></script>
<script type="text/javascript" src="xheditor/xheditor_plugins/ubb.min.js"></script>
<script type="text/javascript">$('#content').xheditor({upLinkUrl:"upload.asp?immediate=1",upLinkExt:"zip,rar,txt",upImgUrl:"upload.asp?immediate=1",upImgExt:"jpg,jpeg,gif,png",upFlashUrl:"upload.asp?immediate=1",upFlashExt:"swf",upFlvUrl:"upload.asp?immediate=1",upFlvExt:"flv",beforeSetSource:ubb2html,beforeGetSource:html2ubb});</script>
<%End Sub%>
<%
'发表
If act="new" And Form_Save="yes" Then
If Title="" Then Call ShowError("标题不能为空！",1)
If Content="" Then Call ShowError("内容不能为空！",1)
If ClassID="" Then Call ShowError("未选择栏目及分类！",1)
If Poster="" Then Poster=User_ID
If TitleBold<>"" Then TitleBold=1 Else TitleBold=0
If Classic<>"" Then Classic=1 Else Classic=0
If SetTop<>"" Then SetTop=1 Else SetTop=0
If Lock<>"" Then Lock=1 Else Lock=0
If Hide<>"" Then Hide=1 Else Hide=0
Set rs=Server.CreateObject("ADODB.Recordset") 
sql="select * from Content"
rs.open sql,conn,3,2
'获取远程图片
If getpicRight="1" Then call myreplace(content)
rs.addnew
rs("Title")=Title
rs("Content") =content
rs("Tags")=Tags
rs("Poster")=Poster
rs("PostTime")=Now
rs("ClassID")=ClassID
rs("subclass")=subclass
Rs("Ip")=RemoteIP()
rs("SetTop")=SetTop
rs("Classic")=Classic
If TitleColor<>"" Then rs("TitleColor")=TitleColor
rs("TitleBold")=TitleBold
rs("Hide")=Hide
rs("Lock")=Lock
rs.update
rs.close
Response.write("<script language='javascript'>alert('发表成功！');window.location='apost.asp';</script>")
End If
'========================================================================
'编辑文章
If act="EditText" And Form_Save="yes" Then
If Title="" Then Call ShowError("标题不能为空！",1)
If Content="" Then Call ShowError("内容不能为空！",1)
If Poster="" Then Poster=User_ID
If TitleBold<>"" Then TitleBold=1 Else TitleBold=0
If Classic<>"" Then Classic=1 Else Classic=0
If SetTop<>"" Then SetTop=1 Else SetTop=0
If Lock<>"" Then Lock=1 Else Lock=0
If Hide<>"" Then Hide=1 Else Hide=0
If Verify<>"" Then Verify=0 Else Verify=-1
sql="select * from Content where id=" &request("artid")
Set rs=Server.CreateObject("ADODB.Recordset") 
rs.open sql,conn,3,3
'获取远程图片
If getpicRight="1" Then call myreplace(content)
rs("Title")=Title
rs("Content") =content
rs("Tags")=Tags
rs("Poster")=Poster
rs("SetTop")=SetTop
rs("Classic")=Classic
rs("TitleColor")=TitleColor
rs("TitleBold")=TitleBold
rs("Lock")=Lock
rs("Hide")=Hide
rs("Verify")=Verify
rs.update
rs.close
Response.write("<script language='javascript'>alert('修改成功！');window.location='admin/admin_show.asp?target=article';</script>")
End If
%>
<%Call CloseAll()%>
