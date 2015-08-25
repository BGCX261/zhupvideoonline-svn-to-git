<!--#include file="UpLoad_Class.asp"-->
<%
dim upload,iserr,id,qid,desc,str
iserr = false
id = 0
qid = 0
desc = ""
set upload = new AnUpLoad
upload.Exe = "jpg|jpeg|gif|png|swf|flv"
upload.MaxSize = 2 * 1024 * 1024 '2M
upload.GetData()
if upload.ErrorID>0 then 
	iserr = true
	desc = upload.Description
else
	dim file,savpath
	savepath = "imgSwf"
	set file = upload.files("file")
	id = upload.forms("pid")
	qid = upload.forms("que")
	if not(file is nothing) then
		set result = file.saveToFile(savepath,0,true)
		if not result.error then
			desc = file.filename
		else
			iserr = true
			desc = file.Exception
		end if
	else
		iserr = true
		desc = "没有上传文件"
	end if
end if
set upload = nothing
%>
<script type="text/javascript">
window.parent.setStep(<%=lcase(iserr)%>,'<%=replace(desc,"'","\'")%>');
</script>