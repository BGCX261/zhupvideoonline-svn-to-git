<%
'**************************************
'**		ofunctions.asp
'**
'** �ļ�˵���������ǳ��ú���
'** �� �� �ţ�v2.2 access gb2312��
'** �޸����ڣ�2008-05-22
'** ��    �ߣ�learn365
'** ��    վ��http://www.learn365.cn
'**************************************

'========================================================================
Function UrlEncode(Str)			'��URL����
	If IsNull(Str) Then
		Str=""
	End If
	UrlEncode=Server.UrlEncode(Str)
End Function
'========================================================================
Function Read(tPath)			'���ļ�
	Dim Fso,Path,File
	Set Fso=CreateObject("Scripting.FileSystemObject")
	If Mid(tPath,2,1)=":" Then
		Path=tPath
	Else
		Path=Server.MapPath(tPath)
	End If
	Set File=Fso.OpenTextFile(Path,1,True)
	If File.AtEndOfStream=False Then
		Read=File.ReadAll
	Else
		Read=""
	End If
	File.Close
	Set File=Nothing
	Set Fso=Nothing
End Function
'========================================================================
Function Save(tPath,Txt)			'д�ļ�
	Dim Fso,Path,File
	Set Fso=CreateObject("Scripting.FileSystemObject")
	If Mid(tPath,2,1)=":" Then
		Path=tPath
	Else
		Path=Server.MapPath(tPath)
	End If
	Set File=Fso.CreateTextFile(Path)
	File.Write(Txt)
	File.Close
	Set File=Nothing
	Set Fso=Nothing
End Function
'========================================================================
Function CCEmpty(StrData)			'�ÿ�
	Dim Str
	Str=StrData
	Str=Replace(Str," ","")
	Str=Replace(Str,"��","")
	CCEmpty=Str
End Function
'========================================================================
Function IsSpecial(StrData)			'�ж��Ƿ���������ַ�
	Dim Str,Fy_In,Fy_Inf,Fy_Xh
	Str=StrData
	Fy_In = "'|;|and|(|)|exec|insert|select|delete|update|count|*|%|$|@|!|(|)|+|=|&|/|\|chr|mid|master|truncate|char|declare" 
	Fy_Inf = split(Fy_In,"|") 
    For Fy_Xh = 0 To Ubound(Fy_Inf)
	If Instr(LCase(Str),Fy_Inf(Fy_Xh)) <> 0 Then
	IsSpecial=Ture
	Exit For
	Else
	IsSpecial=False
	End If 
	Next

'	Num=Len(Str)
'	For i=1 To Num
'		Code=Asc(Mid(Str,i,1))
'		If (Code>=48 And Code<=57) Or (Code>=65 And Code<=90) Or (Code>=97 And Code<=122) Or (Code<0) Then
'			IsSpecial=False
'		Else
'			IsSpecial=True
'			Exit For
'		End If
'	Next
End Function
'========================================================================
function IsValidEmail(email)		'�ж�email�Ƿ�Ϸ�
dim names, name, i, c
'Check for valid syntax in an email address.
IsValidEmail = true
names = Split(email, "@")
if UBound(names) <> 1 then
IsValidEmail = false
exit function
end if
for each name in names
if Len(name) <= 0 then
IsValidEmail = false
exit function
end if
for i = 1 to Len(name)
c = Lcase(Mid(name, i, 1))
if InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 and not IsNumeric(c) then
IsValidEmail = false
exit function
end if
next
if Left(name, 1) = "." or Right(name, 1) = "." then
IsValidEmail = false
exit function
end if
next
if InStr(names(1), ".") <= 0 then
IsValidEmail = false
exit function
end if
i = Len(names(1)) - InStrRev(names(1), ".")
if i <> 2 and i <> 3 then
IsValidEmail = false
exit function
end if
if InStr(email, "..") > 0 then
IsValidEmail = false
end if
end function 
'========================================================================
Function Printl(Str)				'��ʽ���
	If IsNull(Str) Then
		Str=""
	End If
	Response.Write Str & vbcrlf
End Function

'========================================================================
Function Print(Str)					'���
	If IsNull(Str) Then
		Str=""
	End If
	Response.Write Str
End Function
'========================================================================
Function isChinese(para)			'�ж��Ƿ��������
dim str, c
isChinese=false
str=cstr(para)
for i = 1 to Len(para)
c=asc(mid(str,i,1))
If c<0 then
isChinese=true
EXIT Function
End If
Next
End Function
'=======================================================================
'=============================================
'ɾ������
'=============================================

Sub delattach(strUploadFiles)
dim strUploadFile,arrUploadFiles,fso
Set fso = CreateObject("Scripting.FileSystemObject")
if instr(strUploadFiles,"|")>1 then
	arrUploadFiles=split(strUploadFiles,"|")
	for i=0 to ubound(arrUploadFiles)
		if fso.FileExists(server.MapPath("../" & arrUploadfiles(i))) then
			fso.DeleteFile(server.MapPath("../" & arrUploadfiles(i)))
		end if
	next
else
	if fso.FileExists(server.MapPath("../" & strUploadfiles)) then
		fso.DeleteFile(server.MapPath("../" & strUploadfiles))
	end if
end if
Set fso = nothing
End Sub

'=============================================
'ɾ���������ò������ϴ�����
'=============================================
Sub delattachnouse(upfilelist)
dim arrUploadFiles,fso,strRubbishFile
If upfilelist<>"" Then
Set fso = CreateObject("Scripting.FileSystemObject")
   if instr(upfilelist,"|")>1 then
      arrUploadFiles=split(upfilelist,"|")
      upfilelist=""
      for i=0 to ubound(arrUploadFiles)
         if instr(Form_Context,arrUploadfiles(i))<=0 Then
             if fso.FileExists(server.MapPath("" & arrUploadfiles(i))) then
                fso.DeleteFile(server.MapPath("" & arrUploadfiles(i)))
             end If
         else
             if i=0 then
                upfilelist=arrUploadFiles(i)
             else
                upfilelist=upfilelist & "|" & arrUploadFiles(i)
             end if
         end if
      next
   Else
      if instr(Form_Context,upfilelist)<=0 Then
         if fso.FileExists(server.MapPath("" & upfilelist)) then
            fso.DeleteFile(server.MapPath("" & upfilelist))
         end If
         upfilelist=""
      end if
   end if
Set fso = nothing
end if
End Sub


'======================================================

'��������:ReadTextFile
'����:����AdoDb.Stream��������ȡUTF-8��ʽ���ı��ļ�
function ReadFromTextFile (FileUrl,CharSet)
dim str,stm
set stm=server.CreateObject("adodb.stream")
stm.Type=2 '�Ա�ģʽ��ȡ
stm.mode=3 
stm.charset=CharSet
stm.open
stm.loadfromfile server.MapPath(FileUrl)
str=stm.readtext
stm.Close
set stm=nothing
ReadFromTextFile=str
end function
'======================================================
'��������:WriteToTextFile
'����:����AdoDb.Stream������д��UTF-8��ʽ���ı��ļ�
Sub WriteToTextFile (FileUrl,byval Str,CharSet) 
dim stm
set stm=server.CreateObject("adodb.stream")
stm.Type=2 '�Ա�ģʽ��ȡ
stm.mode=3
stm.charset=CharSet
stm.open
stm.WriteText str
stm.SaveToFile server.MapPath(FileUrl),2 
stm.flush
stm.Close
set stm=nothing
end Sub

'*************************************************
'get remote pictures
'*************************************************
Sub myreplace(str)
dim newstr,objregEx,matches,match
newstr=str
set objregEx = new RegExp
objregEx.IgnoreCase = true
objregEx.Global = true
objregEx.Pattern = "http://(.+?)\.(jpg|gif|png|bmp)\[\/img\]" '�����ļ���׺
set matches = objregEx.execute(str)
for each match in matches
saveimg(match)
next
End Sub

Function saveimg(url)
dim savepath,temp,ranNum,filename,xmlhttp,img,objAdostream,SourceFolder,objfso

set objfso=Server.CreateObject("Scripting.FileSystemObject") 
SourceFolder =server.MapPath("UploadFiles/"&DateToStr(Now(),"ymd")&"/") 

'�ж��ļ����Ƿ����,����������򴴽��ļ���
If not objfso.FolderExists(SourceFolder) then
 objfso.CreateFolder SourceFolder    
End if
set objfso=Nothing

savepath="UploadFiles/"&DateToStr(Now(),"ymd")&"/"

url=left(url,len(url)-6)
temp=split(url,".")
'��������ʱ����������������ļ���
randomize
ranNum=int(90000*rnd)+10000
response.write "url="&url
filename=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&"."&temp(ubound(temp))
'�ļ�������������

set xmlhttp=server.createobject("Microsoft.XMLHTTP")
xmlhttp.open "get",url,false
xmlhttp.send
img=xmlhttp.ResponseBody
set xmlhttp=nothing

response.write server.mappath(savepath&filename)

set objAdostream=server.createobject("ADODB.Stream")
objAdostream.Open()
objAdostream.type=1
objAdostream.Write(img)
response.write savepath&filename
objAdostream.SaveToFile(server.mappath(savepath&filename))
objAdostream.SetEOS
set objAdostream=Nothing

' update content
form_context=replace(form_context,url,savepath&filename)
End function
%>
