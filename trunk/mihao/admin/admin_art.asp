<!--#include file="../include/head.asp"-->
<!--#include file="../include/ofunctions.asp"-->
<%
'文件说明：文章管理、浏览页面下的文章操作

Dim ContextID,ParentID,Forum_ID,upfilelist,parentClassid
Act=sqlshow(Request.QueryString("act"))
ContextID=sqlshow(Cnum(Request.QueryString("id")))
If Not IsNumeric(ContextID) Then
	Call ShowError("ID号错误！",1)
End If
Sql="select classid from  Content where ID = " & ContextID 
OpenRs(Sql)
Forum_Id = Rs("classid")
rs.close
If User_Group="admin" Then
Select Case Act
	Case "delete"
		Sql="select * from  Content where ID = " & ContextID & ""
		OpenRs(Sql)
		ParentID = Rs("Parent")
		Rs.Close
		
		If ParentID <> 0 Then  '如果是回复
		  Sql="select replynum,upfilelist from Content where ID=" & ParentID & ""
		   OpenRs(Sql)
	          Rs("ReplyNum")=Rs("ReplyNum")-1
		   Rs.Update
		   Rs.Close
		    '删除附件
		   Sql="select upfilelist from Content where ID=" & ContextID & ""
		   OpenRs(Sql)
		   upfilelist=Rs("upfilelist")
		   Rs.Close
		   if upfilelist<>"" Then
		   delattach(upfilelist)
		   End If
		  Sql="delete * from Content where ID=" & ContextID & ""
		  OpenRs(sql)
		  TurnTo "../context.asp?ID=" & ParentID & ""
		Else                     '主题贴
		   '删除附件
		   Sql="select upfilelist from Content where ID=" & ContextID & " or Parent=" & ContextID
		   OpenRs(Sql)
		   For i=1 to Rs.recordcount
		   upfilelist=Rs("upfilelist")
		   If upfilelist<>"" Then
		   delattach(upfilelist)
		   End If
		   Rs.MoveNext
		   Next
		   Rs.Close
		   Sql="delete * from Content where ID=" & ContextID & " or Parent=" & ContextID
		   OpenRs(Sql)
	       TurnTo "/"
		End If
	Case "deletereply"
		 '删除附件
 		 Sql="select upfilelist from Content where Parent=" & ContextID
    	   OpenRs(Sql)
		   For i=1 to Rs.recordcount
		   upfilelist=Rs("upfilelist")
		   If upfilelist<>"" Then
		   delattach(upfilelist)
		   End If
		   Rs.MoveNext
		   Next
		   Rs.Close
		 Sql="delete * from Content where Parent=" & ContextID
		 OpenRs(sql)
		 Sql="select * from Content where ID=" &ContextID & ""
		 OpenRs(Sql)
	     Rs("ReplyNum")=0
		 Rs("LastUpdateUser")=""
		 Rs.Update
		 TurnTo "../context.asp?ID=" & ContextID & ""
	Case "classic"
		Sql="select * from Content where ID=" & ContextID
		OpenRs(Sql)
		Rs("Classic")=1
		Rs.Update
		Rs.Close
		 TurnTo "../context.asp?ID=" & ContextID & ""
	Case "settop"
		Sql="select * from Content where ID=" & ContextID
		OpenRs(Sql)
		Rs("SetTop")=1
		Rs.Update
		Rs.Close
		 TurnTo "../context.asp?ID=" & ContextID & ""
	Case "settopall"
		Sql="select * from Content where ID=" & ContextID
		OpenRs(Sql)
		Rs("SetTopAll")=1
		Rs.Update
		Rs.Close
		 TurnTo "../context.asp?ID=" & ContextID & ""
	Case "down"
		Sql="update Content set LastUpDateTime=PostTime where ID=" & ContextID
		OpenRs(Sql)
		TurnTo "../context.asp?ID=" & ContextID & ""
	Case "removeclassic"
		Sql="select * from Content where ID=" & ContextID
		OpenRs(Sql)
		Rs("Classic")=0
		Rs.Update
		Rs.Close
		 TurnTo "../context.asp?ID=" & ContextID & ""
	Case "removetop"
		Sql="select * from Content where ID=" & ContextID
		OpenRs(Sql)
		Rs("SetTop")=0
		Rs.Update
		Rs.Close
		 TurnTo "../context.asp?ID=" & ContextID & ""
	Case "removetopall"
		Sql="select * from Content where ID=" & ContextID
		OpenRs(Sql)
		Rs("SetTopAll")=0
		Rs.Update
		Rs.Close
		 TurnTo "../context.asp?ID=" & ContextID & ""
	Case "lock"
		Sql="select * from Content where ID=" & ContextID
		OpenRs(Sql)
		Rs("lock")=true
		Rs.Update
		Rs.Close
		 TurnTo "../context.asp?ID=" & ContextID & ""
	Case "unlock"
		Sql="select * from Content where ID=" & ContextID
		OpenRs(Sql)
		Rs("lock")=false
		Rs.Update
		Rs.Close
		 TurnTo "../context.asp?ID=" & ContextID & ""
	Case "hide"
	    Sql="select * from Content where ID=" & ContextID
		OpenRs(Sql)
		Rs("hide")=True
		Rs.Update
		ParentID = Rs("Parent")
		Rs.Close
		If ParentID = 0 Then
		TurnTo "../context.asp?id=" & ContextID
		Else 
		TurnTo "../context.asp?id=" & ParentID
		End If
	Case "unhide"
	    Sql="select * from Content where ID=" & ContextID
		OpenRs(Sql)
		Rs("hide")=false
		Rs.Update
		ParentID = Rs("Parent")
		If ParentID = 0 Then
		TurnTo "../context.asp?id=" & ContextID
		Else 
		TurnTo "../context.asp?id=" & ParentID
		End If	
	Case "setbold"
		sql="update content set titleBold=1 where ID=" & ContextID & ""
		openrs sql
		 TurnTo "../context.asp?ID=" & ContextID & ""
	Case "resetbold"
		sql="update content set titleBold=0 where ID=" & ContextID & ""
		openrs sql
		 TurnTo "../context.asp?ID=" & ContextID & ""
	Case "color"
		Sql="select * from Content where ID=" & ContextID
		OpenRs(Sql)
		Rs("TitleColor")=true
		Rs.Update
		Rs.Close
		 TurnTo "../context.asp?ID=" & ContextID & ""
	Case "tcolor"
		Sql="select * from Content where ID=" & ContextID
		OpenRs(Sql)
		Rs("TitleColor")=false
		Rs.Update
		Rs.Close
		 TurnTo "../context.asp?ID=" & ContextID & ""
	Case "class"
		Dim ClassIndex
		ClassIndex=Cnum(Request.QueryString("classindex"))+1
		Sql="update Content set ClassID="&ClassIndex&" where ID=" & ContextID & " or parent=" & ContextID
		OpenRs(Sql)
		Sql="update Content set subclass=0 where ID=" & ContextID
		OpenRs(Sql)
		 TurnTo "../context.asp?ID=" & ContextID & ""
	Case "subclass"
		ClassIndex=Cnum(Request.QueryString("subclassid"))
		parentClassid=sqlshow(Cnum(Request.QueryString("parentClassid")))
		Sql="update Content set ClassID="&parentClassid&" where ID=" & ContextID  & " or parent=" & ContextID

		OpenRs(Sql)
		Sql="update Content set subclass="&ClassIndex&" where ID=" & ContextID
		OpenRs(Sql)
		 TurnTo "../context.asp?ID=" & ContextID & ""
	Case "toTopic"
		dim t_title,t_context,t_user,t_text,newparentID,subclass,newTitle
		newTitle=Request.QueryString("title")
		Sql="select * from Content where ID="&ContextID&""
		OpenRs(Sql)
		t_title=Rs("title")
		t_context=Rs("Content")
		t_user=Rs("Poster")
		newparentID=Rs("Parent")
		subclass=Rs("subclass")
		Rs.Close
		Sql="select subclass from Content where ID="&newparentID&""
		OpenRs(Sql)
		subclass=Rs("subclass")
		Rs.Close
		t_text="[quote2][color=gray]此主题于"&Now&"被"&User_ID&"由下面主题的回复信息生成:[/color]"&CHR(10)&"[url=./"&strUrl2&"?id="&newparentID&"]"&t_title&"[/url][/quote2]"&CHR(10)
		Sql="select * from Content where id=null"
		OpenRs(Sql)
		Rs.AddNew
		Rs("Title")="[转入主题] "&newTitle
		Rs("Content")=t_text&t_Context
		Rs("Poster")=t_user
		Rs("Parent")=0
		Rs("PostTime")=Now
		Rs("classID")=Forum_ID
		Rs("LastUpdateTime")=Now
		Rs("Ip")=RemoteIP()
		Rs("subclass")=subclass
		Rs.Update
		Rs.Close
		Response.write("<script language=""javascript"">window.opener.document.location.reload();window.close();</script>")
End Select
CloseAll
Else
	TurnTo strUrl
End If

%>