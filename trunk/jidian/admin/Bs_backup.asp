<!--#include file="bsconfig.asp"-->
<!-- #include file="Inc/Head.asp" -->
<%call checkmanage("05")%>
<table align="center" class="a2">
	<tr>
		<td class="a1">�������ݿ�</td>
	</tr>
	<tr class="a4">
		<td>
<%
if request("action")="Backup" then
call backupdata()
else
%>
			<form method="post" action="Bs_backup.asp?action=Backup">
				<table width="100%">
					<tr> 
						<td colspan="2" class="iptsubmit">�����̳������ļ�[��ҪFSOȨ��]</td>
					</tr>
					<tr> 
						<td class="iptlab"> ��ǰ���ݿ�·��</td>
						<td class="ipt"><input type=text size=50 name=DBpath value="<%=db%>"></td>
					</tr>
					<tr> 
						<td class="iptlab"> �������ݿ�Ŀ¼[��Ŀ¼�����ڣ������Զ�����]</td>
						<td class="ipt"><input type=text size=50 name=bkfolder value=Databackup></td>
					</tr>
					<tr> 
						<td class="iptlab">�������ݿ�����[�籸��Ŀ¼�и��ļ��������ǣ���û�У����Զ�����]</td>
						<td class="ipt"><input type=text size=30 name=bkDBname value=ebossi_data_backup.mdb></td>
					</tr>
					<tr> 
						<td colspan="2" class="iptsubmit">
						<input type=submit value="ȷ��">
						</td>
					</tr>
					<tr> 
						<td colspan="2">
						<br>
						�������Ĭ�����ݿ��ļ�Ϊ<%=db%><br>
						������������������������ķ������ݣ��Ա�֤�������ݰ�ȫ��<br>
						ע�⣺����·��������������ռ��Ŀ¼�����·��</td>
					</tr>
				</table>
			</form>
      <%end if%><% 
sub backupdata() 
Dbpath=request.form("Dbpath") 
Dbpath=server.mappath(Dbpath) 
bkfolder=request.form("bkfolder") 
bkdbname=request.form("bkdbname") 
Set Fso=server.createobject("scripting.filesystemobject") 
if fso.fileexists(dbpath) then 
If CheckDir(bkfolder) = True Then 
fso.copyfile dbpath,bkfolder& "\"& bkdbname 
else 
MakeNewsDir bkfolder 
fso.copyfile dbpath,bkfolder& "\"& bkdbname 
end if 
response.write "�������ݿ�ɹ��������ݵ����ݿ�·��Ϊ" &bkfolder& "\"& bkdbname 
Else 
response.write "�Ҳ���������Ҫ���ݵ��ļ���" 
End if 
end sub 
'------------------���ĳһĿ¼�Ƿ����------------------- 
Function CheckDir(FolderPath) 
folderpath=Server.MapPath(".")&"\"&folderpath 
Set fso1 = CreateObject("Scripting.FileSystemObject") 
If fso1.FolderExists(FolderPath) then 
'���� 
CheckDir = True 
Else 
'������ 
CheckDir = False 
End if 
Set fso1 = nothing 
End Function 
'-------------����ָ����������Ŀ¼--------- 
Function MakeNewsDir(foldername) 
Set fso1 = CreateObject("Scripting.FileSystemObject") 
Set f = fso1.CreateFolder(foldername) 
MakeNewsDir = True 
Set fso1 = nothing 
End Function 
%>
		</td>
  </tr>
</table>
<%htmlend%>
