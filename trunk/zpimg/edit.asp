<!--#include file="include/head.asp"-->
<!--#include file="include/ofunctions.asp"-->
<%
Dim Form_ID,vl_code,Form_Context,Form_SubClass,Form_Title,Form_TitleBold,Form_TitleColor,getpicRight
Act=Request.QueryString("act")
Form_ID=Request.QueryString("id")
vl_code=trim(request.form("vl_code"))

If Act="" Then Turnto "ssb.asp"

If User_Group<>"admin" Then
If cstr(vl_code)="" or cstr(Session("GetCode"))<>cstr(vl_code) Then Call ShowError("��������ȷ����֤��!",1)
End If

Form_Context=Request.Form("Content")
Form_SubClass=Cnum(Request.Form("subclass"))
Form_Title=Request.Form("Title")
Form_TitleBold=Request.Form("TitleBold")
Form_TitleColor=Request.Form("TitleColor")
forumID=Cnum(Request.Form("forumid"))
Form_ID=Cnum(Request.Form("artid"))
getpicRight=Cnum(Request.Form("getpicRight"))

If Form_Context="" Then Call ShowError("���ݲ���Ϊ�գ�",1)
If Form_TitleBold<>"" Then Form_TitleBold=1 Else Form_TitleBold=0

Select Case Act

'�ظ�����
Case "reply"
Sql="select * from Content where ID=" & SqlShow(Form_ID)
OpenRs(Sql)
'������������
Rs("LastUpdateTime")=Now
Rs("ReplyNum")=Rs("ReplyNum")+1
Rs.Update
'�����ظ�����
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
'�����û���������
Sql="select * from Users where UserName='"&SqlShow(User_ID)&"'"
OpenRs Sql
Rs("ReplyNum")=Rs("ReplyNum")+1
Rs("Bonus")=Rs("Bonus")+2
Rs.Update
Rs.Close
'�ظ���ɷ�������
Turnto "context.asp?ID="&Form_ID&""

'ɾ����������
Case "delete"
sql="delete * from Content where id=" &request("id")
OpenRs(sql)
Application(SystemKey&"Loaded")=""
Call ShowError("ɾ���ɹ���",1)

'Ͷ��
Case "tougao" 
Form_Poster=Request.Form("Poster")
If Form_Title="" Then  Call ShowError("����д���⣡",1)
If Form_Poster="" Then  Call ShowError("����д��Դ��Ϣ��",1)
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
If instr(Form_Context,"[img]") > 0 Then Rs("isPic")=1 '����ͼƬ��ʶ
Rs.Update
Rs.Close
'�����û���������
Sql="select * from Users where UserName='"&SqlShow(User_ID)&"'"
OpenRs Sql
Rs("PostNum")=Rs("PostNum")+1
Rs.Update
Rs.Close
'Ͷ�����
Response.write("<script language='javascript'>alert('Ͷ��ɹ�����ȴ���ˣ�лл����Mihao.net��֧�֣�');window.location='ssb.asp';</script>")

Case Else
End Select
%>
