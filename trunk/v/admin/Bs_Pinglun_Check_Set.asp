<!--#include file="bsconfig.asp"-->
<!--#include file="../Inc/Config.asp"-->
<!--#include file="Inc/Function.asp"-->

<%
dim PinglunID,ActionH,sqlDel,rsDel,FoundErr,ErrMsg,ObjInstalled

PinglunID=trim(request("PinglunID"))
ActionH=Trim(Request("ActionH"))
FoundErr=False
ObjInstalled=IsObjInstalled("Scripting.FileSystemObject")

if PinglunID="" or ActionH="" then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
end if
if FoundErr=False then
	if instr(PinglunID,",")>0 then
		dim idarr,i
		idArr=split(PinglunID)
		for i = 0 to ubound(idArr)
			call CheckPinglun(clng(idarr(i)),ActionH)
		next
	else
		call CheckPinglun(clng(PinglunID),ActionH)
	end if
end if
if FoundErr=False then
	call CloseConn()
	response.Redirect "Bs_Pinglun_Check.asp"
else
	call CloseConn()
	call WriteErrMsg()
end if

sub CheckPinglun(ID,CheckAction)
	sqlDel="select * from buyok_pinglun where id=" & CLng(ID)
	Set rsDel= Server.CreateObject("ADODB.Recordset")
	rsDel.open sqlDel,conn,1,3
	if rsDel.bof and rsDel.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到：" & rsDel("id") & " </li>"
	end if
		if CheckAction="OkCheck" then
			rsDel("shenhe")=True
		elseif CheckAction="CancelCheck" then
			rsDel("shenhe")=False
		end if
		rsDel.update
		set rsDel=nothing
end sub
%>