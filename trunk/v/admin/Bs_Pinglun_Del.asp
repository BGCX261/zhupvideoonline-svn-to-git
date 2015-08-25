<!--#include file="bsconfig.asp"-->
<!--#include file="Inc/Function.asp"-->
<!--#include file="../Inc/Config.asp"-->

<%
dim PinglunID,Act,sqlDel,rsDel,FoundErr,ErrMsg,ObjInstalled,prodid
PinglunID=trim(request("PinglunID"))
prodid=trim(request("prodid"))
'PinglunID=replace(request("PinglunID")," ","")
Act=Trim(Request("Act"))
FoundErr=False
ObjInstalled=IsObjInstalled("Scripting.FileSystemObject")

if PinglunID="" or Act<>"Del" then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
end if
if FoundErr=False then
	if instr(PinglunID,",")>0 and instr(prodid,",")>0  then
		dim idarr,i,proarr,p
		idArr=split(PinglunID)
		proarr=split(prodid)
		for i = 0 to ubound(idArr) 
			call DelArticle(clng(idarr(i)))
		next
		for p = 0 to ubound(proarr)
		    call DelVot(clng(proarr(p)))
		next
	else
		call DelArticle(clng(PinglunID))
		call DelVot(clng(prodid))
	end if
end if
if FoundErr=False then
	call CloseConn()
	response.Redirect "Bs_Pinglun_Check.asp"
else
	call CloseConn()
	call WriteErrMsg()
end if

'删除评论
sub DelArticle(ID)
	sqlDel="select * from buyok_pinglun where id=" & CLng(ID)
	Set rsDel= Server.CreateObject("ADODB.Recordset")
	rsDel.open sqlDel,conn,1,3
	if FoundErr=False then
		rsDel.delete
		rsDel.update
		set rsDel=nothing
	end if
end sub

'删除评论时减去相应人气
sub DelVot(Pro)
	set rsV=Server.CreateObject("ADODB.RecordSet")
	sqlV="select * from Product where ArticleID=" & CLng(Pro)
	rsV.open sqlV,conn,1,3
	rsV("Votes")=rsV("Votes")-1
	rsV.Update
    rsV.close
	set rsV=nothing
	'conn.execute "update Product set Votes =Votes-1 where ArticleID=" &  CLng(Pro)
end sub
%>