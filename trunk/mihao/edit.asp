<!--#include file="include/head.asp"-->
<!--#include file="include/ofunctions.asp"-->
<%
Dim Form_ID,vl_code,Form_Context,Form_SubClass,Form_Title,Form_TitleBold,Form_TitleColor,getpicRight
Act=Request.QueryString("act")
Form_ID=Request.QueryString("id")
vl_code=trim(request.form("vl_code"))

If Act="" Then Turnto "ssb.asp"

If User_Group<>"admin" Then
If cstr(vl_code)="" or cstr(Session("GetCode"))<>cstr(vl_code) Then Call ShowError("请输入正确的验证码!",1)
End If

Form_Context=Request.Form("Content")
Form_SubClass=Cnum(Request.Form("subclass"))
Form_Title=Request.Form("Title")
Form_TitleBold=Request.Form("TitleBold")
Form_TitleColor=Request.Form("TitleColor")
forumID=Cnum(Request.Form("forumid"))
Form_ID=Cnum(Request.Form("artid"))
getpicRight=Cnum(Request.Form("getpicRight"))

If Form_Context="" Then Call ShowError("内容不能为空！",1)
If Form_TitleBold<>"" Then Form_TitleBold=1 Else Form_TitleBold=0

Select Case Act

'回复文章
Case "reply"
Sql="select * from Content where ID=" & SqlShow(Form_ID)
OpenRs(Sql)
'更新所属文章
Rs("LastUpdateTime")=Now
Rs("ReplyNum")=Rs("ReplyNum")+1
Rs.Update
'创建回复内容
Rs.AddNew
Rs("Parent")=Form_ID
Rs("Content")=Form_Context
Rs("Poster")=User_ID
Rs("PostTime")=Now
Rs("LastUpdateTime")=Now
Rs("Ip")=RemoteIP()
Rs("ClassID")=1
Rs.Update
Rs.Close
'更新用户发帖数量
Sql="select * from Users where UserName='"&SqlShow(User_ID)&"'"
OpenRs Sql
Rs("ReplyNum")=Rs("ReplyNum")+1
Rs("Bonus")=Rs("Bonus")+2
Rs.Update
Rs.Close
'回复完成返回文章
Turnto "context.asp?ID="&Form_ID&""

'删除文章评论
Case "delete"
sql="delete * from Content where id=" &request("id")
OpenRs(sql)
Application(SystemKey&"Loaded")=""
Call ShowError("删除成功！",1)

'投稿
Case "tougao" 
Form_Poster=Request.Form("Poster")
If Form_Title="" Then  Call ShowError("请填写标题！",1)
If Form_Poster="" Then  Call ShowError("请填写来源信息！",1)
Sql="select * from Content"
OpenRs(Sql)
Rs.AddNew
Rs("Title")=Form_Title
Rs("Content")=Form_Context
Rs("Poster")=Form_Poster
Rs("ClassID")=999
Rs("Tags")=User_ID
Rs("PostTime")=Now
Rs("Ip")=RemoteIP()
Rs("Parent")=0
Rs("YesNo")=1
If instr(Form_Context,"[img]") > 0 Then Rs("isPic")=1 '加入图片标识
Rs.Update
Rs.Close
'更新用户发帖数量
Sql="select * from Users where UserName='"&SqlShow(User_ID)&"'"
OpenRs Sql
Rs("PostNum")=Rs("PostNum")+1
Rs.Update
Rs.Close
'投稿完成
Response.write("<script language='javascript'>alert('投稿成功，请等待审核！谢谢您对Mihao.net的支持！');window.location='ssb.asp';</script>")

Case Else
End Select
%>
