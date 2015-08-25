<%
'-----------------------------------------------------
' ����: Asp������ϴ���������
' ����: ����(http://blog.joycode.com/dotey)
' ����: ����ʵ����(http://www.openlab.net.cn), ������(http://blog.joycode.com), ����԰(http://www.cnblogs.com), ���ǽű�(http://www.51js.com)
' �汾: 1.0
' ��Ȩ: ����Ʒ�����ʹ�ã����������Ƴ���Ȩ��Ϣ
' ����: http://www.webuc.net/MyProject/upload/upload.rar
' �Ƽ�: China Community Server(http://www.CommunityServer.cn/)��һ����Դ��Asp.Net����ϵͳ��������̳��Blog���������԰�ȳ���
'-----------------------------------------------------

Dim DoteyUpload_SourceData

Class DoteyUpload
	
	Public Files
	Public Form
	Public MaxTotalBytes
	Public Version
	Public ProgressID
	Public ErrMsg
	
	Private BytesRead
	Private ChunkReadSize
	Private Info
	Private Progress

	Private UploadProgressInfo
	Private CrLf

	Private Sub Class_Initialize()
		Set Files = Server.CreateObject("Scripting.Dictionary")	' �ϴ��ļ�����
		Set Form = Server.CreateObject("Scripting.Dictionary")	' ��������
		UploadProgressInfo = "DoteyUploadProgressInfo"  ' Application��Key
		MaxTotalBytes = 60 *1024 *1024 ' Ĭ�����60��
		ChunkReadSize = 64 * 1024	' �ֿ��С64K
		CrLf = Chr(13) & Chr(10)	' ����

		Set DoteyUpload_SourceData = Server.CreateObject("ADODB.Stream")
		DoteyUpload_SourceData.Type = 1 ' ��������
		DoteyUpload_SourceData.Open

		Version = "1.0 Beta"	' �汾
		ErrMsg = ""	' ������Ϣ
		Set Progress = New ProgressInfo

	End Sub

	' ���ļ��������ļ���ͳһ������ĳ·����
	Public Sub SaveTo(path)
		
		Upload()	' �ϴ�

		if right(path,1) <> "/" then path = path & "/" 

		' �����������ϴ��ļ�
		For Each fileItem In Files.Items			
			fileItem.SaveAs path & fileItem.FileName
		Next

		' �����������½�����Ϣ
		Progress.ReadyState = "complete" '�ϴ�����
		UpdateProgressInfo progressID

	End Sub

	' �����ϴ������ݣ������浽��Ӧ������
	Public Sub Upload ()

		Dim TotalBytes, Boundary
		TotalBytes = Request.TotalBytes	 ' �ܴ�С
		If TotalBytes < 1 Then
			Raise("�����ݴ���")
			Exit Sub
		End If
		If TotalBytes > MaxTotalBytes Then
			Raise("����ǰ�ϴ���СΪ" & TotalBytes/1000 & " K���������Ϊ" & MaxTotalBytes/1024 & "K")
			Exit Sub
		End If
		Boundary = GetBoundary()
		If IsNull(Boundary) Then 
			Raise("���form��û�а���multipart/form-data�ϴ�����Ч��")
			Exit Sub	 ''���form��û�а���multipart/form-data�ϴ�����Ч��
		End If
		Boundary = StringToBinary(Boundary)
		
		Progress.ReadyState = "loading" '��ʼ�ϴ�
		Progress.TotalBytes = TotalBytes
		UpdateProgressInfo progressID

		Dim DataPart, PartSize
		BytesRead = 0

		'ѭ���ֿ��ȡ
		Do While BytesRead < TotalBytes

			'�ֿ��ȡ
			PartSize = ChunkReadSize
			if PartSize + BytesRead > TotalBytes Then PartSize = TotalBytes - BytesRead
			DataPart = Request.BinaryRead(PartSize)
			BytesRead = BytesRead + PartSize

			DoteyUpload_SourceData.Write DataPart

			Progress.UploadedBytes = BytesRead
			Progress.LastActivity = Now()

			' ���½�����Ϣ
			UpdateProgressInfo progressID

		Loop

		' �ϴ���������½�����Ϣ
		Progress.ReadyState = "loaded" '�ϴ�����
		UpdateProgressInfo progressID

		Dim Binary
		DoteyUpload_SourceData.Position = 0
		Binary = DoteyUpload_SourceData.Read

		Dim BoundaryStart, BoundaryEnd, PosEndOfHeader, IsBoundaryEnd
		Dim Header, bFieldContent
		Dim FieldName
		Dim File
		Dim TwoCharsAfterEndBoundary

		BoundaryStart = InStrB(Binary, Boundary)
		BoundaryEnd = InStrB(BoundaryStart + LenB(Boundary), Binary, Boundary, 0)

		Do While (BoundaryStart > 0 And BoundaryEnd > 0 And Not IsBoundaryEnd)
			' ��ȡ����ͷ�Ľ���λ��
			PosEndOfHeader = InStrB(BoundaryStart + LenB(Boundary), Binary, StringToBinary(vbCrLf + vbCrLf))
						
			' �������ͷ��Ϣ�������ڣ�
			' Content-Disposition: form-data; name="file1"; filename="G:\homepage.txt"
			' Content-Type: text/plain
			Header = BinaryToString(MidB(Binary, BoundaryStart + LenB(Boundary) + 2, PosEndOfHeader - BoundaryStart - LenB(Boundary) - 2))

			' �����������
			bFieldContent = MidB(Binary, (PosEndOfHeader + 4), BoundaryEnd - (PosEndOfHeader + 4) - 2)
			
			FieldName = GetFieldName(Header)
			' ����Ǹ���
			If InStr (Header,"filename=""") > 0 Then
				Set File = New FileInfo
				
				' ��ȡ�ļ������Ϣ
				Dim clientPath
				clientPath = GetFileName(Header)
				File.FileName = GetFileNameByPath(clientPath)
				File.FileExt = GetFileExt(clientPath)
				File.FilePath = clientPath
				File.FileType = GetFileType(Header)
				File.FileStart = PosEndOfHeader + 3
				File.FileSize = BoundaryEnd - (PosEndOfHeader + 4) - 2
				File.FormName = FieldName

				' ������ļ���Ϊ�ղ������ڸñ������֮
				If Not Files.Exists(FieldName) And File.FileSize > 0 Then
					Files.Add FieldName, File
				End If
			'��������				
			Else
				' ����ͬ������
				If Form.Exists(FieldName) Then
					Form(FieldName) = Form(FieldName) & "," & BinaryToString(bFieldContent)
				Else
					Form.Add FieldName, BinaryToString(bFieldContent)
				End If
			End If

			' �Ƿ����λ��
			TwoCharsAfterEndBoundary = BinaryToString(MidB(Binary, BoundaryEnd + LenB(Boundary), 2))
			IsBoundaryEnd = TwoCharsAfterEndBoundary = "--"

			If Not IsBoundaryEnd Then ' ������ǽ�β, ������ȡ��һ��
				BoundaryStart = BoundaryEnd
				BoundaryEnd = InStrB(BoundaryStart + LenB(Boundary), Binary, Boundary)
			End If
		Loop
		
		' �����ļ���������½�����Ϣ
		Progress.UploadedBytes = TotalBytes
		Progress.ReadyState = "interactive" '�����ļ�����
		UpdateProgressInfo progressID

	End Sub

	'�쳣��Ϣ
	Private Sub Raise(Message)
		ErrMsg = ErrMsg & "[" & Now & "]" & Message & "<BR>"
		
		Progress.ErrorMessage = Message
		UpdateProgressInfo ProgressID
		
		'call Err.Raise(vbObjectError, "DoteyUpload", Message)

	End Sub

	' ȡ�߽�ֵ
	Private Function GetBoundary()
		Dim ContentType, ctArray, bArray
		ContentType = Request.ServerVariables("HTTP_CONTENT_TYPE")
		ctArray = Split(ContentType, ";")
		If Trim(ctArray(0)) = "multipart/form-data" Then
			bArray = Split(Trim(ctArray(1)), "=")
			GetBoundary = "--" & Trim(bArray(1))
		Else	'���form��û�а���multipart/form-data�ϴ�����Ч��
			GetBoundary = null
			Raise("���form��û�а���multipart/form-data�ϴ�����Ч��")
		End If
	End Function

	' ����������ת�����ı�
	Private Function BinaryToString(xBinary)
		Dim Binary
		if vartype(xBinary) = 8 then Binary = MultiByteToBinary(xBinary) else Binary = xBinary
		
	  Dim RS, LBinary
	  Const adLongVarChar = 201
	  Set RS = CreateObject("ADODB.Recordset")
	  LBinary = LenB(Binary)
		
		if LBinary>0 then
			RS.Fields.Append "mBinary", adLongVarChar, LBinary
			RS.Open
			RS.AddNew
				RS("mBinary").AppendChunk Binary 
			RS.Update
			BinaryToString = RS("mBinary")
		Else
			BinaryToString = ""
		End If
	End Function


	Function MultiByteToBinary(MultiByte)
	  Dim RS, LMultiByte, Binary
	  Const adLongVarBinary = 205
	  Set RS = CreateObject("ADODB.Recordset")
	  LMultiByte = LenB(MultiByte)
		if LMultiByte>0 then
			RS.Fields.Append "mBinary", adLongVarBinary, LMultiByte
			RS.Open
			RS.AddNew
				RS("mBinary").AppendChunk MultiByte & ChrB(0)
			RS.Update
			Binary = RS("mBinary").GetChunk(LMultiByte)
		End If
	  MultiByteToBinary = Binary
	End Function


	' �ַ�����������
	Function StringToBinary(String)
		Dim I, B
		For I=1 to len(String)
			B = B & ChrB(Asc(Mid(String,I,1)))
		Next
		StringToBinary = B
	End Function

	'���ر�����
	Private Function GetFieldName(infoStr)
		Dim sPos, EndPos
		sPos = InStr(infoStr, "name=")
		EndPos = InStr(sPos + 6, infoStr, Chr(34) & ";")
		If EndPos = 0 Then
			EndPos = inStr(sPos + 6, infoStr, Chr(34))
		End If
		GetFieldName = Mid(infoStr, sPos + 6, endPos - _
			(sPos + 6))
	End Function

	'�����ļ���
	Private Function GetFileName(infoStr)
		Dim sPos, EndPos
		sPos = InStr(infoStr, "filename=")
		EndPos = InStr(infoStr, Chr(34) & CrLf)
		GetFileName = Mid(infoStr, sPos + 10, EndPos - _
			(sPos + 10))
	End Function

	'�����ļ��� MIME type
	Private Function GetFileType(infoStr)
		sPos = InStr(infoStr, "Content-Type: ")
		GetFileType = Mid(infoStr, sPos + 14)
	End Function

	'����·����ȡ�ļ���
	Private Function GetFileNameByPath(FullPath)
		Dim pos
		pos = 0
		FullPath = Replace(FullPath, "/", "\")
		pos = InStrRev(FullPath, "\") + 1
		If (pos > 0) Then
			GetFileNameByPath = Mid(FullPath, pos)
		Else
			GetFileNameByPath = FullPath
		End If
	End Function

	'����·����ȡ��չ��
	Private Function GetFileExt(FullPath)
		Dim pos
		pos = InStrRev(FullPath,".")
		if pos>0 then GetFileExt = Mid(FullPath, Pos)
	End Function

	' ���½�����Ϣ
	' ������Ϣ������Application�е�ADODB.Recordset������
	Private Sub UpdateProgressInfo(progressID)
		Const adTypeText = 2, adDate = 7, adUnsignedInt = 19, adVarChar = 200
		
		If (progressID <> "" And IsNumeric(progressID)) Then
			Application.Lock()
			if IsEmpty(Application(UploadProgressInfo)) Then
				Set Info = Server.CreateObject("ADODB.Recordset")
				Set Application(UploadProgressInfo) = Info
				Info.Fields.Append "ProgressID", adUnsignedInt
				Info.Fields.Append "StartTime", adDate
				Info.Fields.Append "LastActivity", adDate
				Info.Fields.Append "TotalBytes", adUnsignedInt
				Info.Fields.Append "UploadedBytes", adUnsignedInt
				Info.Fields.Append "ReadyState", adVarChar, 128
				Info.Fields.Append "ErrorMessage", adVarChar, 4000
				Info.Open 
		 		Info("ProgressID").Properties("Optimize") = true
				Info.AddNew 
			Else
				Set Info = Application(UploadProgressInfo)
				If Not Info.Eof Then
					Info.MoveFirst()
					Info.Find "ProgressID = " & progressID
				End If
				If (Info.EOF) Then
					Info.AddNew
				End If
			End If

			Info("ProgressID") = clng(progressID)
			Info("StartTime") = Progress.StartTime
			Info("LastActivity") = Now()
			Info("TotalBytes") = Progress.TotalBytes
			Info("UploadedBytes") = Progress.UploadedBytes
			Info("ReadyState") = Progress.ReadyState
			Info("ErrorMessage") = Progress.ErrorMessage
			Info.Update

			Application.UnLock
		End IF
	End Sub

	' �����ϴ�ID��ȡ������Ϣ
	Public Function GetProgressInfo(progressID)

		Dim pi, Infos
		Set pi = New ProgressInfo
		If Not IsEmpty(Application(UploadProgressInfo)) Then
			Set Infos = Application(UploadProgressInfo)
			If Not Infos.Eof Then
				Infos.MoveFirst
				Infos.Find "ProgressID = " & progressID
				If Not Infos.EOF Then
					pi.StartTime = Infos("StartTime")
					pi.LastActivity = Infos("LastActivity")
					pi.TotalBytes = clng(Infos("TotalBytes"))
					pi.UploadedBytes = clng(Infos("UploadedBytes"))
					pi.ReadyState = Trim(Infos("ReadyState"))
					pi.ErrorMessage = Trim(Infos("ErrorMessage"))
					Set GetProgressInfo = pi
				End If
			End If
		End If
		Set GetProgressInfo = pi
	End Function

	' �Ƴ�ָ���Ľ�����Ϣ
	Private Sub RemoveProgressInfo(progressID)
		If Not IsEmpty(Application(UploadProgressInfo)) Then
			Application.Lock
			Set Info = Application(UploadProgressInfo)
			If Not Info.Eof Then
				Info.MoveFirst
				Info.Find "ProgressID = " & progressID
				If  Not Info.EOF Then
					Info.Delete
				End If
			End If

			' ���û�м�¼��, ֱ���ͷ�, ����'800a0bcd'����
			If Info.RecordCount = 0 Then
				Info.Close
				Application.Contents.Remove UploadProgressInfo
			End If
			Application.UnLock
		End If
	End Sub

	' �Ƴ�ָ���Ľ�����Ϣ
	Private Sub RemoveOldProgressInfo(progressID)
		If Not IsEmpty(Application(UploadProgressInfo)) Then
			Dim L
			Application.Lock

			Set Info = Application(UploadProgressInfo)
			Info.MoveFirst

			Do
				L = Info("LastActivity").Value
				If IsEmpty(L) Then
					Info.Delete() 
				ElseIf DateDiff("d", Now(), L) > 30 Then
					Info.Delete()
				End If
				Info.MoveNext()
			Loop Until Info.EOF

			' ���û�м�¼��, ֱ���ͷ�, ����'800a0bcd'����
			If Info.RecordCount = 0 Then
				Info.Close
				Application.Contents.Remove UploadProgressInfo
			End If
			Application.UnLock
		End If
	End Sub

End Class

'---------------------------------------------------
' ������Ϣ ��
'---------------------------------------------------
Class ProgressInfo
	
	Public UploadedBytes
	Public TotalBytes
	Public StartTime
	Public LastActivity
	Public ReadyState
	Public ErrorMessage

	Private Sub Class_Initialize()
		UploadedBytes = 0	' ���ϴ���С
		TotalBytes = 0	' �ܴ�С
		StartTime = Now()	' ��ʼʱ��
		LastActivity = Now()	 ' ������ʱ��
		ReadyState = "uninitialized"	' uninitialized,loading,loaded,interactive,complete
		ErrorMessage = ""
	End Sub

	' �ܴ�С
	Public Property Get TotalSize
		TotalSize = FormatNumber(TotalBytes / 1024, 0, 0, 0, -1) & " K"
	End Property 

	' ���ϴ���С
	Public Property Get SizeCompleted
		SizeCompleted = FormatNumber(UploadedBytes / 1024, 0, 0, 0, -1) & " K"
	End Property 

	' ���ϴ�����
	Public Property Get ElapsedSeconds
		ElapsedSeconds = DateDiff("s", StartTime, Now())
	End Property 

	' ���ϴ�ʱ��
	Public Property Get ElapsedTime
		If ElapsedSeconds > 3600 then
			ElapsedTime = ElapsedSeconds \ 3600 & " ʱ " & (ElapsedSeconds mod 3600) \ 60 & " �� " & ElapsedSeconds mod 60 & " ��"
		ElseIf ElapsedSeconds > 60 then
			ElapsedTime = ElapsedSeconds \ 60 & " �� " & ElapsedSeconds mod 60 & " ��"
		else
			ElapsedTime = ElapsedSeconds mod 60 & " ��"
		End If
	End Property 

	' ��������
	Public Property Get TransferRate
		If ElapsedSeconds > 0 Then
			TransferRate = FormatNumber(UploadedBytes / 1024 / ElapsedSeconds, 2, 0, 0, -1) & " K/��"
		Else
			TransferRate = "0 K/��"
		End If
	End Property 

	' ��ɰٷֱ�
	Public Property Get Percentage
		If TotalBytes > 0 Then
			Percentage = fix(UploadedBytes / TotalBytes * 100) & "%"
		Else
			Percentage = "0%"
		End If
	End Property 

	' ����ʣ��ʱ��
	Public Property Get TimeLeft
		If UploadedBytes > 0 Then
			SecondsLeft = fix(ElapsedSeconds * (TotalBytes / UploadedBytes - 1))
			If SecondsLeft > 3600 then
				TimeLeft = SecondsLeft \ 3600 & " ʱ " & (SecondsLeft mod 3600) \ 60 & " �� " & SecondsLeft mod 60 & " ��"
			ElseIf SecondsLeft > 60 then
				TimeLeft = SecondsLeft \ 60 & " �� " & SecondsLeft mod 60 & " ��"
			else
				TimeLeft = SecondsLeft mod 60 & " ��"
			End If
		Else
			TimeLeft = "δ֪"
		End If
	End Property 

End Class

'---------------------------------------------------
' �ļ���Ϣ ��
'---------------------------------------------------
Class FileInfo
	
	Dim FormName, FileName, FilePath, FileSize, FileType, FileStart, FileExt, NewFileName

	Private Sub Class_Initialize 
		FileName = ""		' �ļ���
		FilePath = ""			' �ͻ���·��
		FileSize = 0			' �ļ���С
		FileStart= 0			' �ļ���ʼλ��
		FormName = ""	' ������
		FileType = ""		' �ļ�Content Type
		FileExt = ""			' �ļ���չ��
		NewFileName = ""	'�ϴ����ļ���
	End Sub

	Public Function Save()
		SaveAs(FileName)
	End Function

	' �����ļ�
	Public Function SaveAs(fullpath)
		Dim dr
		SaveAs = false
		If trim(fullpath) = "" Or FileStart = 0 Or FileName = "" Or right(fullpath,1) = "/" Then Exit Function
		
		NewFileName = GetFileNameByPath(fullpath)

		Set dr = CreateObject("Adodb.Stream")
		dr.Mode = 3
		dr.Type = 1
		dr.Open
		DoteyUpload_SourceData.position = FileStart
		DoteyUpload_SourceData.copyto dr, FileSize
		dr.SaveToFile MapPath(FullPath), 2
		dr.Close
		set dr = nothing 
		SaveAs = true
	End function

	' ����Binary
	Public Function GetBinary()
		Dim Binary
		If FileStart = 0 Then Exit Function

		DoteyUpload_SourceData.Position = FileStart
		Binary = DoteyUpload_SourceData.Read(FileSize)

		GetBinary = Binary
	End function

	' ȡ��������·��
	Private Function MapPath(Path)
		If InStr(1, Path, ":") > 0 Or Left(Path, 2) = "\\" Then
			MapPath = Path 
		Else 
			MapPath = Server.MapPath(Path)
		End If
	End function

	'����·����ȡ�ļ���
	Private Function GetFileNameByPath(FullPath)
		Dim pos
		pos = 0
		FullPath = Replace(FullPath, "/", "\")
		pos = InStrRev(FullPath, "\") + 1
		If (pos > 0) Then
			GetFileNameByPath = Mid(FullPath, pos)
		Else
			GetFileNameByPath = FullPath
		End If
	End Function

End Class

%>