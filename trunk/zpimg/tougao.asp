<!--#include file="include/head.asp"-->
<!--#include file="include/ofunctions.asp"-->
<%
Dim Form_ID,Form_context,Title,Content,poster,tags,ClassID,SubClass,SetTop,Classic,TitleBold,TitleColor,Lock,Hide,Form_Save,j,getpicRight
Act=Request.QueryString("act")
Form_ID=Request.QueryString("id")
Form_Save=Request.Form("save")
Title=Request.Form("Title")
Content=Request.Form("Content")
Poster=Request.Form("Poster")
ClassID=Request.Form("ForumID")
SubClass=Request.Form("SubClass")
SetTop=Request.Form("SetTop")
Classic=Request.Form("Classic")
getpicRight=Cnum(Request.Form("getpicRight"))

If User_ID="" Then 
Call ShowError("您尚未登录！",1)
End If 

If act="" Then
act="new"
Call PostEdit
End If

Sub HtmlTltle '浏览器标题
Response.write strSiteName
End Sub

Sub HtmlKeywords '关键字
Response.write strKeyword
End Sub

Sub HtmlDescription '描述
Response.write strAbout
End Sub

Sub PostEdit
%>
<!--#include file="html/Header.html"-->
<div id="positionbg">
    <div id="position">当前位置：<a href="/">首页</a>&nbsp;&gt;&nbsp;会员投稿</div>
</div>
<div id="TouGao">
<form name="form1" method="POST" action="tougao.asp?act=<%=act%>">
<input type="hidden" name="save" value="yes">
<span style="display:none"><a id='fff'><%Application(systemkey&ForumID)%></a></span>
<p>文章标题 <input type="text" name="title" class="TGInput1" maxlength="60">&nbsp;&nbsp;&nbsp;&nbsp;<%
Response.write "隶属栏目 <select class=""input"" name=""forumID"" id=""classname"" onchange=""javascript:setclass();"" >"
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
Response.write "&nbsp;<select class=""button"" name=""subclass"" id=""subclass"" ><option value="""">选择分类</option>"
For j=1 to ubound(subclassArray)
Response.write("<option ")
If j=Cnum(subclassID) Then response.write " selected "
Response.write("value="&j&">"&subclassArray(j)&"</option>")
Next
Response.write "</select> <span class=""red"">*请注意选择隶属栏目及二级分类</span>"
%>
</p>
<div class="divtext">
<textarea id="content" name="content" class="xheditor" rows="12" cols="80" style="width:1250px;height:300px;"></textarea>
</div>
<p><input type="submit" value="确定发表" class="TGbtn"></p>
</div>
<!--#include file="html/Footer.html"-->

<script language="javascript" >
//分类自动更新
function setclass()
{
document.all("fff").innerHTML="<font color=red><b>发贴目标栏目已更改</b></font>";
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
<script type="text/javascript">$('#content').xheditor({beforeSetSource:ubb2html,beforeGetSource:html2ubb});</script>
<%End Sub%>
<%
'发表
If act="new" And Form_Save="yes" Then
If Title="" Then Call ShowError("标题不能为空！",1)
If Content="" Then Call ShowError("内容不能为空！",1)
If ClassID="" Then Call ShowError("未选择栏目及分类！",1)
Set rs=Server.CreateObject("ADODB.Recordset") 
sql="select * from Content"
rs.open sql,conn,3,2
rs.addnew
rs("Title")=Title
rs("Content") =content
rs("Poster")=User_ID
rs("PostTime")=Now
rs("ClassID")=ClassID
rs("subclass")=subclass
rs("Ip")=RemoteIP()
rs("Submission")=-1
rs("Verify")=-1
rs.update
rs.close

'邮件提醒
If strSubTsEmail="1" Then
	Dim theMail
	Set theMail=Server.CreateObject("CDONTS.NewMail")
	theMail.From="china@china.com"
	theMail.To =strAdminEmail
	theMail.Subject ="您的网站有新的投稿"
	theMail.BodyFormat=0
	theMail.MailFormat=0
	theMail.Body="投稿会员："&User_ID&"<br>文章标题："&Title&""
	theMail.Send
	set theMail=nothing 
End If

Response.write("<script language='javascript'>alert('投稿成功，请等待管理员审核！');window.location='tougao.asp';</script>")
End If
%>
<%Call CloseAll()%>
