<!--#include file="../include/head.asp"-->
<!--#include file="../include/ofunctions.asp"-->
<%
'�ļ�˵������̨������ʾģ��
If User_Group<>"admin" Then Call ShowError("��ҳֻ�й���Ա���ܷ��ʣ�",1)
dim bgcolor,m,ii,rs2,sql2,rs3,sql3,target,page,Pages,PageSize,AllCount,ShowCount,PastCount,sqlstr,j,orderstr,wherestr
target=Request.QueryString("target")
if target="" then target="info"
Page=Max(Cnum(Request("page")),1)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/css.css" rel="stylesheet" type="text/css">
<script language="javascript" src="../include/edit.js"></script>
<script language="javascript">

function askdelete()
{
	if(confirm("�Ƿ�ȷ��ɾ����"))
		return(true);
	else
		return(false);
}

function checkform()
{
	if(form1.boardid.value==""){
		alert(" ��ĿID����Ϊ�գ�");
		form1.boardid.focus();
		return(false);
	}
	if(form1.boardname.value==""){
		alert("��Ŀ���Ʋ���Ϊ�գ�");
		form1.boardname.focus();
		return(false);
	}
	return(true);
}


function unselectall()
{
    if(document.myform.chkAll.checked){
	document.myform.chkAll.checked = document.myform.chkAll.checked&0;
    } 	
}

function CheckAll(form)
{
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.Name != "chkAll")
       e.checked = form.chkAll.checked;
    }
}

function setclass()
{
var boardCount=<%=BoardCount%>
var subclass='<%=Application(SystemKey&"subclassAll")%>';
var pro_Class=new Array();
var pro_SubClass=subclass.split(',');

pro_Class[0]='--����--';
pro_Class[boardCount+1]='<%=Application(Systemkey&"subclass999")%>';

for(z=1;z<=boardCount;z++) {
pro_Class[z]=pro_SubClass[z-1];
}

var nowSelectIndex=document.all("classid").selectedIndex+1;
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
</head>

<body>

<%Select Case target%> <%
'===========================================
'������Ϣ
     Case "info"
     dim sitename%>
<div class="DB">
    <div class="L1">��ǰλ�ã�������� &gt;</div>
	<div class="D1">
		 <div class="T1">�ҵĲ�����Ϣ</div>
		 <ul class="U1">
		 <%
		 Sql="select * from Users where UserName='"&User_ID&"'"
		 OpenRs(Sql)
		 Response.write("<li>�ϴε�¼ʱ�䣺"&Rs("LastLogonTime")&"</li>")
		 Response.write("<li>�ϴε�¼IP��"&Rs("Ip")&"</li>")
		 Rs.Close
		 %>
			  <li class="border-top"><a href="?target=SetLog">���������־</a>��</li>
		 <%
		 Sql="select * from SetLog where AdminName='"&User_ID&"' order by id desc"
		 OpenRs(Sql)
		 If Rs.RecordCount>3 Then ii=3 Else ii=Rs.RecordCount
		 For i=1 To ii
		 Response.write("<li>"&Rs("SetTxt")&"��"&Rs("SetTime")&"��</li>")
		 Rs.Movenext
		 Next
		 Rs.Close
		 %>
	     </ul>
    </div>
	<div class="G1">&nbsp;</div>
	<div class="D1">
		 <div class="T1">��ݲ���</div>
		 <ul class="U2">
		      <li><a href="admin_show.asp?target=ShengCheng">������վ��ҳ</a></li>
		      <li><a href="javascript:opennotice('ks')">����������Ϣ</a></li>
			  <li><a href="javascript:openhotkw('kw')">���������ȴ�</a></li>
			  <li><a href="javascript:openjifen('')">�޸Ļ�Ա����</a></li>
			  <li><a href="admin_all.asp?target=remove">������վ����</a></li>
			  <li><a href="http://www.baidu.com/s?wd=site:<%=request.ServerVariables("HTTP_HOST")%>" target="_blank">�ٶ���¼���</a></li>
	     </ul>
    </div>
	<div class="G2">&nbsp;</div>
	<div class="D1">
		 <div class="T1"><a href="admin_show.asp?target=Submission">����Ͷ��</a></div>
		 <ul class="U1">
		<%sql="select * from Content where Submission=-1 order by Verify asc,id desc"
    openrs(sql)
If Rs.RecordCount>6 Then ii=6 Else ii=Rs.RecordCount
For i=1 To ii
Response.write("<li>["&Rs("Poster")&"] <a href=""../context.asp?id="&rs("id")&""" target=""_blank"">"&Rs("Title")&"</a>")
If Rs("Verify")=-1 Then
Response.write(" <font color=""#FF0000"">δ���</font>")
%>
��<a href="admin_all.asp?target=article&act=deletetougao&id=<%=rs("id")%>" onClick="javascript:return confirm('ȷ��ɾ��������?(��������һ�лظ�)');">ɾ��</a> - <a href="../apost.asp?act=edit&id=<%=rs("id")%>">�༭</a>��
<%
End If
Response.write("</li>")
Rs.Movenext
Next
Rs.Close
%>
	     </ul>
    </div>
	<div class="G1">&nbsp;</div>
	<div class="D1">
		 <div class="T1">�׺Źٷ���̬</div>
		 <ul class="U1"><iframe src="http://net.mihao.net/notice/index.html" frameborder="0" scrolling="no" height="160px" width="350px" noresize="noresize"></iframe></ul>
    </div>
</div>
<%

Function GetTotalSize(GetLocal,GetType)
	Dim FSO
	Set FSO=Server.CreateObject("Scripting.FileSystemObject")
	IF Err<>0 Then
		Err.Clear
		GetTotalSize="�������ر�FSO���鿴ռ�ÿռ�ʧ��"
	Else
		Dim SiteFolder
		IF GetType="Folder" Then
			Set SiteFolder=FSO.GetFolder(GetLocal) 
		Else
			Set SiteFolder=FSO.GetFile(GetLocal) 
		End IF
		GetTotalSize=SiteFolder.Size
		IF GetTotalSize>1024*1024 Then
		GetTotalSize=GetTotalSize/1024/1024
		IF inStr(GetTotalSize,".") Then GetTotalSize = Left(GetTotalSize,inStr(GetTotalSize,".")+2)
			GetTotalSize=GetTotalSize&" MB"
		Else
			GetTotalSize=Fix(GetTotalSize/1024)&" KB"
		End IF
		Set SiteFolder=Nothing
	End IF
	Set FSO=Nothing
End Function

'==========================================
'��վ����
	Case "setup"
	dim txt
	txt=Read("../include/config.asp")
	%>
<p align="center"><font color="#FF0000"><b>��վ��������</b></font> </p>
<form name="forumsetup" method="POST" action="admin_all.asp?target=setup&act=save"></form>
<table cellspacing="1" cellpadding="5" border="0" class="a23" bgcolor="#FFFFFF" align="center" width="631">
	<tr class="a1">
		<td>������վ��������</td>
	</tr>
	<tr class="a4">
		<td><textarea rows="28" name="setup" cols="85"><%=txt%></textarea></td>
	</tr>
	<tr class="a1">
		<td height="95">
		<p align="center">
		<input type="submit" value="��������" name="savesetup" class="button"></p>
		</td>
	</tr>
</table>
<%'==========================================
'��Ŀ����
     Case "board"%>
<div class="DB">
    <div class="L1">��ǰλ�ã�<a href="admin_show.asp">�������</a> &gt; ��Ŀ����</div>
    <form name="form1" method="POST" action="admin_all.asp?target=board&act=new" onSubmit="javascript:return checkform();">
	<table cellspacing="1" cellpadding="5" border="0" class="a23" bgcolor="#FFFFFF" align="center">
		<tr class="BoardTitle">
			<td width="30" class="center bold">ID</td>
			<td width="150" class="center bold">��Ŀ����</td>
			<td width="150" class="center bold">��Ŀ����</td>
			<td width="70" class="center bold">��������</td>
			<td width="150" class="center bold">����</td>
		</tr>
		<%sql="select * from Board where BoardID<>999 order by BoardID"
    openrs(sql)
	For i=1 To Rs.RecordCount
	If i mod 2 = 0 Then 
          bgcolor = "UsersTr1"
        Else
         bgcolor = "UsersTr2" 
     End If
     %>
		<tr onMouseOver="this.className='ahover'" onMouseOut="this.className='<%=bgcolor%>'" class="<%=bgcolor%>">
			<td class="center"><%=rs("boardid")%></td>
			<td class="center"><%=rs("boardname")%></td>
			<td class="center">
<%
If rs("BoardType")=0 Then Response.Write("����")
If rs("BoardType")=1 Then Response.Write("ͼƬ")
If rs("BoardType")=2 Then Response.Write("��Ƶ")
%></td>
			<td class="center"><font style="font-size: 8pt"><%'��ʾ�ð�������Ϣ
		sql2="select count(*) from content where classid="&rs("boardid")&"and parent=0"
        Set Rs2=Server.CreateObject("Adodb.Recordset")
        Rs2.Open sql2,Conn,1,3
        Response.write rs2(0)
        rs2.close

%></font> </td>
			<td class="center">
			<a href="admin_all.asp?target=board&act=edit&boardid=<%=rs("boardid")%>">
			�޸�</a> <span lang="en-us">&nbsp; </span><%If rs("boardid")<>999 Then%><a href="admin_all.asp?target=board&act=delete&boardid=<%=rs("boardid")%>" onClick="javascript:return askdelete();"> 
			ɾ��</a> <%End If%>
			<a href="admin_all.asp?target=board&act=clear&boardid=<%=rs("boardid")%>" onClick="javascript:return confirm('ȷ����ձ���Ŀ�𣿣�����');">&nbsp;���</a>
			</td>
		</tr>
		<%
	rs.movenext
	Next
	%>
		<tr class="UsersTr2">
			<td class="center"><input type="hidden" name="save" value="yes">
			<input type="text" name="boardid" size="2" maxlength="4" class="input"></td>
			<td class="center"><input type="text" name="boardname" maxlength="20" class="IT2"></td>
			<td class="center">
			<select class="Sel" name="BoardType" id="BoardType">
			<option value="0">ѡ������</option>
			<option value="0">����</option>
			<option value="1">ͼƬ</option>
            <option value="2">��Ƶ</option>
			</select>
			</td>
			<td class="center"></td>
			<td class="center" colspan="3"><span lang="en-us">&nbsp;<input type="submit" value="ȷ��" name="addforum" class="btn2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<a href="admin_all.asp?target=board&act=lineup">
			<font color="#FF0000">��������</font></a></span></td>
		</tr>
	</table>
</form>
<form name="form2" method="POST" action="admin_all.asp?target=board&act=move">
	<p class="BoardFoot">ת��<select class="Sel" name="sourceboard" id="sourceboard">
	<option value="998">ѡ����Ŀ</option> <%for i=1 to BoardCount
	Response.write "<option value='"&i&"'>"&Application(systemkey&i)&"</option>"
	Next
	%></option>
	</select> �������µ� <select class="Sel" name="targetboard" id="targetboard">
	<option value="998">ѡ����Ŀ</option> <%for i=1 to BoardCount
Response.write "<option value='"&i&"'>"&Application(systemkey&i)&"</option>"
Next
%></option>
	</select>&nbsp;&nbsp;<input type="submit" name="Submit1" value="ִ��" class="btn2"></p>
</form>
<p class="BoardFoot2">ע�⣺��ĿID�����1��ʼ���������������֣��粻�������Ե㡰�������򡱡���Ŀһ���趨�����к�������ʱ�벻Ҫ���׸Ķ���ĿID��</p>
</div>
<%
'==========================================
'������Ŀ����
 Case "subclass"
 %>
<div class="DB">
    <div class="L1">��ǰλ�ã�<a href="admin_show.asp">�������</a> &gt; �������</div>
<table cellspacing="1" cellpadding="5" border="0" class="a23" bgcolor="#FFFFFF" align="center" width="631">
	<tr class="T3">
		<td width="29" class="center bold">ID</td>
		<td width="195" class="center bold">��������</td>
		<td width="89" class="center bold">����/�ظ�</td>
		<td width="122" class="center bold">����Ŀ����</td>
		<td width="134" class="center bold">����</td>
	</tr>
	<%FOR j=1 to boardcount
		If j mod 2 = 0 Then 
        bgcolor = "ClassBg2"
        Else
         bgcolor = "ClassBg1" 
	    End If
%>
	<tr>
		<td colspan="5" class="<%=bgcolor%>"><b><%=Application(SystemKey&j)%></b></td>
	</tr>
	<%sql="select * from subclass where subparent="&j&" order by subnum"
	    openrs(sql)
		If Rs.RecordCount>0 Then
		For i=1 To Rs.RecordCount
	     %>
		<tr onMouseOver="this.className='ahover'" onMouseOut="this.className='<%=bgcolor%>'" class="<%=bgcolor%>">
		<td width="29" class="center"><%=rs("subnum")%></td>
		<td width="195" class="center"><%=rs("subname")%></td>
		<td width="89" class="center"><%'��ʾ�÷�����Ϣ
		sql2="select count(*) from content where classid="&rs("subparent")&" and parent=0 and subclass="&Rs("subnum")&""
        Set Rs2=Server.CreateObject("Adodb.Recordset")
        Rs2.Open sql2,Conn,1,3
        Response.write rs2(0)
        rs2.close
        
        sql2="select sum(replynum) from content where classid="&rs("subparent")&"and parent=0 and subclass="&Rs("subnum")&""
        Set Rs2=Server.CreateObject("Adodb.Recordset")
        Rs2.Open sql2,Conn,1,3
        Response.write "/"&rs2(0)
        rs2.close

%></td>
		<td width="122" class="center"><%=Application(SystemKey&rs("subparent"))%></td>
		<td width="134" class="center">
		<a href="admin_all.asp?target=subclass&act=edit&boardid=<%=rs("subparent")%>&subnum=<%=Rs("subnum")%>">
		�޸�</a> <span lang="en-us">&nbsp; </span>
		<a href="admin_all.asp?target=subclass&act=delete&boardid=<%=rs("subparent")%>&subnum=<%=Rs("subnum")%>" onClick="javascript:return askdelete();">
		ɾ��</a>&nbsp; </td>
	</tr>
	<%
rs.movenext
Next
Else
%>
	<tr class="<%=bgcolor%>">
		<td colspan="5" class="center">����Ŀ���޷��࣬���½����࣡</td>
	</tr>
	<%End If
Rs.Close
%>
	<form name="form<%=j%>" method="POST" action="admin_all.asp?target=subclass&act=new">
		<tr class="<%=bgcolor%>">
			<td width="29" class="center">
			<input type="hidden" name="save" value="yes">
			<input type="hidden" name="subparent" value="<%=j%>">
			<input type="text" name="subid" size="3" maxlength="4" class="IT3"></td>
			<td width="195" class="center">
			<input type="text" name="subname" maxlength="20" class="IT2">
			</td>
			<td colspan="3" class="center"><input type="submit" value="ȷ�����" name="addforum" class="btn3">&nbsp;&nbsp;<a href="admin_all.asp?target=subclass&act=lineup&subparent=<%=j%>">
			<font color="#FF0000">��������</font></a></td>
		</tr>
	</form>
	<form name="form_move" method="POST" action="admin_all.asp?target=subclass&act=move">
		<tr class="<%=bgcolor%>">
			<td colspan="5">
			<p align="center">ת��<select class="Sel" name="sourceclass" id="sourceclass">
			<%
	If instr(Application(SystemKey&"subclass"&j),"|")>1 Then
	subclassArray=split(Application(SystemKey&"subclass"&j),"|")
	Else
	subclassArray=split("����","|")
	End If
	for i=0 to ubound(subclassArray)
	Response.write "<option value='"&i&"'>"&subclassArray(i)&"</option>"
	Next
	%></select> �����е� <select class="Sel" name="targetclass" id="targetclass">
			<%for i=0 to ubound(subclassArray)
Response.write "<option value='"&i&"'>"&subclassArray(i)&"</option>"
Next
%></select> <input type="submit" name="Submit1" value="ִ��" class="btn2">
			</p>
			</td>
		</tr>
		<input type="hidden" name="subparent" value="<%=j%>">
	</form>
	<%
NEXT%>
</table>
<p class="div2">ע�⣺����ID��ʾ���Ƿ����˳���粻�������Ե㡰�������򡱡�����һ���趨�����к�������ʱ�벻Ҫ���׸Ķ�����ID��</p>
</form>
</div>
<%
'=========================================================
'��Ϣ����
Case "notice"
dim xx,SetStyle
xx=Request.QueryString("xx")

If xx="ks" Then
SetStyle="style=""display:none"""
End If
%>
<div class="DB">
<div class="L1" <%=SetStyle%>>��ǰλ�ã�<a href="admin_show.asp">�������</a> &gt; ��Ϣ����</div>
<form name="form1" method="POST" action="admin_all.asp?target=notice&act=new" >
<table cellspacing="1" cellpadding="5" border="0" class="a23" bgcolor="#FFFFFF" align="center">
<tr class="T4">
<td width="50" class="center bold">����</td>
<td width="275" class="center bold">����</td>
<td width="275" class="center bold">���ӵ�ַ</td>
<td width="100" class="center bold">�Ƿ���ʾ</td>
<td width="100" class="center bold">����</td>
</tr>
<%
Sql="select * from notice order by OrderID DESC"
Openrs(sql)

PageSize=15
AllCount=Rs.RecordCount
Pages=Fix(AllCount/PageSize)
If Pages*PageSize<AllCount Then
Pages=Pages+1	
End If
PastCount=(Page-1)*PageSize
If PastCount>=AllCount Then
Page=1
PastCount=0
End If
If AllCount>PastCount Then
Rs.Move PastCount
End If
ShowCount=Min(AllCount-PastCount,PageSize)

For i=1 To ShowCount
If i mod 2 = 0 Then 
bgcolor = "Tr1"
Else
bgcolor = "Tr2" 
End If
%>
<tr onMouseOver="this.className='ahover'" onMouseOut="this.className='<%=bgcolor%>'" class="<%=bgcolor%>" <%=SetStyle%>>
<td class="center"><%=rs("OrderID")%></td>
<td>&nbsp;&nbsp;<a href="../<%=rs("Nurl")%>" target="_blank" title="��Ϣ����ʱ�䣺<%=rs("NPostTime")%>"><%=rs("NTitle")%></a></td>
<td>&nbsp;&nbsp;<a href="../<%=rs("Nurl")%>" target="_blank"><%=rs("Nurl")%></a></td>
<td class="center"><%If Rs("EnableNotice")=True Then%>
<a href="admin_all.asp?target=notice&act=disable&id=<%=rs("id")%>">��</a></p>
<%Else%>
<a href="admin_all.asp?target=notice&act=enable&id=<%=rs("id")%>">��</a>
<%End If%>
</td>
<td class="center"><a href="admin_all.asp?target=notice&act=edit&id=<%=rs("id")%>">�޸�</a> - <a href="admin_all.asp?target=notice&act=delete&id=<%=rs("id")%>" onClick="javascript:return askdelete();">ɾ��</a></td>
</tr>
<%
Rs.Movenext
Next
%>
<tr class="Tr3">
<td class="center"><input type="text" name="OrderID" class="IT3"></td>
<td>&nbsp;&nbsp;<input type="text" name="NTitle" class="IT5"></td>
<td>&nbsp;&nbsp;<input type="text" name="Nurl" class="IT5"></td>
<td class="center">
<input type="checkbox" name="EnableNotice" value="ON" checked></td>
<td class="center"><input type="submit" value="ȷ�����" name="addurl" class="btn3"></td>
</tr>
<%
Rs.Close
%>
</table>

<p class="Pages3" <%=SetStyle%>>
<span>��ʾ��������ֵԽ����ʾλ��Խ��ǰ�������������ӵ�ַ�ɲ鿴��������</span>
<%If page>1 Then%>
<a href="?target=notice&page=<%=page-1%>">��һҳ</a>&nbsp;
<%End If%>
<%
Response.write ShowPages(Pages,Page,"?target=notice")
%>
<%If page<pages Then%>
&nbsp;<a href="?target=notice&page=<%=page+1%>">��һҳ</a>
<%End If%>
</p>
</div>
<%
'========================================================
'�û�����
Case "users"%> <a name="#up"></a>
<div class="DB">
    <div class="L1">��ǰλ�ã�<a href="admin_show.asp">�������</a> &gt; ��Ա����</div>
<form name="myform" method="POST" action="admin_all.asp?target=users">
	<table cellspacing="1" cellpadding="5" border="0" class="a23" bgcolor="#FFFFFF" align="center" width="631">
		<tr class="UsersTitle">
			<td width="50" class="center">��</td>
			<td width="104" class="center bold"><a href="admin_show.asp?target=users&sqlstr=name">�û���</a></td>
			<td width="69" class="center bold"><a href="admin_show.asp?target=users&sqlstr=usergroup">�û���</a></td>
			<td width="69" class="center bold"><a href="admin_show.asp?target=users&sqlstr=regtime">ע��ʱ��</a></td>
			<td width="33" class="center bold"><a href="admin_show.asp?target=users&sqlstr=postcount">����</a></td>
			<td width="110" class="center bold"><a href="admin_show.asp?target=users&sqlstr=logintime">��¼ʱ��</a>/<a href="admin_show.asp?target=users&sqlstr=logincount">����</a></td>
			<td width="37" class="center bold">״̬</td>
			<td width="100" class="center bold">����</td>
		</tr>
		<%
    sqlstr=Request.Querystring("sqlstr")
    Select Case sqlstr
    case "name" 
    	sql="select * from users order by username"
    case "regtime"
    	sql="select * from users order by regtime desc"
    case "postcount" 
    	sql="select * from users order by postnum desc"
    case "logintime" 
    	sql="select * from users order by lastlogontime desc"
    case "logincount"
    	sql="select * from users order by logoncount desc"
    case "usergroup" 
    	sql="select * from users order by usergroup"
    case "admin" 
    	sql="select * from users where usergroup='admin' order by username"
    case "master"
    	sql="select * from users where usergroup<>'admin' and usergroup<>'member' order by username"
    case "vip"
    	sql="select * from users where vip=True order by username"
    case "lock" 
		sql="select * from users where lockuser=True order by username"
    case "search" 
    	sql="select * from users where username like '%"&Request("searchname")&"%' order by regtime"
    	dim searchusername
    	searchusername=Request("searchname")
	case else
    	sql="select * from users order by regtime desc"
    End Select

    openrs(sql)
    PageSize=15
    AllCount=Rs.RecordCount
	Pages=Fix(AllCount/PageSize)
	If Pages*PageSize<AllCount Then
		Pages=Pages+1	
	End If
	PastCount=(Page-1)*PageSize
	If PastCount>=AllCount Then
		Page=1
		PastCount=0
	End If
	If AllCount>PastCount Then
		Rs.Move PastCount
	End If
	ShowCount=Min(AllCount-PastCount,PageSize)
	%>
		<input type="hidden" name="page" value="<%=page%>"><%For i=1 To ShowCount
	If i mod 2 = 0 Then 
         bgcolor = "UsersTr1"
       Else
         bgcolor = "UsersTr2" 
     End If
     %>
		<tr onMouseOver="this.className='ahover'" onMouseOut="this.className='<%=bgcolor%>'" class="<%=bgcolor%>">
			<td class="center">
			<input name="UserID" type="checkbox" onClick="unselectall()" id="UserID" value="<%=rs("UserName")%>"></td>
			<td width="104" class="center">
			<a href="javascript:memberinfo2('<%=rs("username")%>')"><%
		Response.write Rs("UserName")
		Response.write "</font></a>"
%></a></td>
			<td width="69" class="center"><%
			
			If Cnum(Rs("usergroup"))>0 Then 
				Response.write "<font color=red>"&Application(systemkey&Cnum(rs("usergroup")))&"</font>"  
				Else 
				Response.write Rs("usergroup") 
			End If
			If Rs("vip")=True Then
			Response.write "<br><font color=red>*VIP</font>"
			End If%> </td>
			<td width="69" class="center"><%=rs("regtime")%></td>
			<td width="33" class="center"><%=rs("PostNum")%></td>
			<td width="158" class="center"><%=rs("lastlogontime")%>[<font color="red"><%=rs("logoncount")%></font>]</td>
			<td class="center" width="29"><%if rs("lockuser")=TRUE then Response.write "<font color=red>����</font>" else Response.write "<font color=green>����</font>"%></b></td>
			<td class="center">
			<a href="javascript:openjifen('<%=rs("UserName")%>')">����</a> - <a href="admin_all.asp?target=users&act=delete&username=<%=rs("UserName")%>" onClick="javascript:return confirm('ȷ��ɾ�����û�?(��������һ�кۼ�)');">ɾ��</a> - <%if rs("lockuser")=False Then
      Response.write "<a href=""admin_all.asp?target=users&sqlstr="&sqlstr&"&act=lock&page="&page&"&username="&rs("UserName")&""">����</a></td></tr>" 
      else 
      Response.write "<a href=""admin_all.asp?target=users&sqlstr="&sqlstr&"&act=unlock&page="&page&"&username="&rs("UserName")&""">����</a></td></tr>"
      End If
      rs.movenext
    Next%> </td>
		</tr>
		<tr class="a10">
			<td class="center">
			<input name="chkAll" type="checkbox" id="chkAll" onClick="CheckAll(this.form)" value="checkbox"></td>
			<td colspan="7" class="center">��������
			<input name="Action" type="radio" value="delete">ɾ��&nbsp;
			<input name="Action" type="radio" value="unlock">����&nbsp;
			<input name="Action" type="radio" value="lock">����&nbsp;
			<input name="Action" type="radio" value="initpass">��������&nbsp;
			<input name="Action" type="radio" value="vip">��ΪVIP&nbsp;
			<input name="Action" type="radio" value="removevip">���VIP&nbsp;
			<input name="Action" type="radio" value="move">�ƶ��� 
			<select class="Sel" name="UserLevel" id="UserLevel">
			<option value="member">��ͨ��Ա</option>
			<option value="admin">������Ա</option>
			<%
                rs.close%></select>&nbsp;&nbsp;
			<input type="submit" name="Submit" value="ִ��" class="btn2">
			</td>
		</tr>
	</table>
</form>
<%'��ʾע���û�����
    sql="select count(*) from users"
    openrs sql%>
<p class="UsersUp">
<span>����λ<font color="red"><%=rs(0)%></font>ע���û�</span>
<a href="admin_show.asp?target=users&page=<%=page-1%>&sqlstr=<%=sqlstr%>">��һҳ</a>&nbsp;<%
Response.write ShowPages(Pages,Page,"admin_show.asp?target=users&sqlstr="&sqlstr&"")
%>&nbsp;<a href="admin_show.asp?target=users&page=<%=page+1%>&sqlstr=<%=sqlstr%>&searchname=<%=searchusername%>">��һҳ</a>
</p>
<form method="POST" action="admin_show.asp?target=users&sqlstr=search#up">
<p class="div2"><a href="admin_show.asp?target=users&sqlstr=">�����û��б�</a> | <a href="admin_show.asp?target=users&sqlstr=admin#up">����Ա�б�</a> | <a href="admin_show.asp?target=users&sqlstr=vip#up">VIP��Ա�б�</a> | <a href="admin_show.asp?target=users&sqlstr=lock#up">�����û��б�</a>&nbsp;&nbsp;<input type="text" name="searchname" size="20" class="IT2">&nbsp;<input type="submit" value="����" name="B1" class="btn2"></p>
</form>
<p class="div2">��ʾ����������ƿ��ԶԸ��н������򣻵���û����ɲ鿴�û�Ա�ĸ������ϡ�</p>
<%
'========================================================
'���¹���
	Case "article"
	forumid=Request.QueryString("forumid")
	sqlstr=Request.Querystring("sqlstr")
	subclassid=Request.QueryString("subclassid")
	'��ȡ����Ŀ
	tempClass=Application(SystemKey&"subclass"&ForumID)
	If instr(tempClass,"|")>0 Then
	subclassArray=split(tempClass,"|")
	Else
	subclassArray=split("����","|")
	End If
	if forumid="" then
        If sqlstr="replynum" Then orderStr=" order by replynum desc"
    	If sqlstr="clicknum" Then orderStr=" order by clicknum desc"
		If sqlstr="" Then orderStr=" order by settop desc,PostTime desc"
		whereStr=" where parent=0" 
	Else
    	If sqlstr="replynum" Then orderStr= " order by replynum desc"
    	If sqlstr="clicknum" Then orderStr= " order by clicknum desc"
    	If sqlstr="classid" Or sqlstr="" Then orderStr= "  order by PostTime desc"
		whereStr=" where parent=0 and classid="&forumid&"" 
    End If
	If IsNumeric(subclassID)=True and Cnum(subclassID)>0 and Cnum(subclassID)<999 Then
	whereStr=whereStr&" and subclass="&subclassID&""
	End If
	If Cnum(subclassID)=999 Then
	whereStr=whereStr&" and subclass<1"
	End If
	sql="select * from content "&whereStr&orderStr
	openrs(sql)
		%>
<div class="DB">
    <div class="L1">��ǰλ�ã�<a href="admin_show.asp">�������</a> &gt; ���¹���</div>
    <form name="myform" method="POST" action="admin_all.asp?target=article">
	<div class="D3">
		 <div class="T2">
		 <%If forumID="" Then%>
		 <a href="admin_show.asp?target=article">ȫ������</a> <%For j=1 To BoardCount
				Print "| <a href=admin_show.asp?target=article&sqlstr=classid&forumID=" &j&  ">"&Application(systemkey&j)&"</a> "
			Next%>
	<%Else
	Response.write("&nbsp;<a href=admin_show.asp?target=article>ȫ����Ŀ</a>&nbsp;|&nbsp;")
	for i=1 to ubound(subclassArray)
	If i=cnum(subclassID) Then 
	Response.write("&nbsp;-&nbsp;<b>"&subclassArray(i)&"</b>")
	Else
	response.write("&nbsp;-&nbsp;<a class='classnametitle' href=admin_show.asp?target=article&forumID="&forumID&"&subclassID="&i&">"&subclassArray(i)&"</a>")
	End If
	next
	response.write("&nbsp;-&nbsp;<a class='classnametitle' href=admin_show.asp?target=article&forumID="&forumID&"&subclassID=999>δ����</a>")
	End If%>
	</div>
	<table cellspacing="0" cellpadding="4" border="0" class="a23" bgcolor="#FFFFFF" align="center">
		<% 
	'�趨ÿҳ��ʾ������
	PageSize=18
    AllCount=Rs.RecordCount
	Pages=Fix(AllCount/PageSize)
	If Pages*PageSize<AllCount Then
		Pages=Pages+1	
	End If
	PastCount=(Page-1)*PageSize
	If PastCount>=AllCount Then
		Page=1
		PastCount=0
	End If
	If AllCount>PastCount Then
		Rs.Move PastCount
	End If
	ShowCount=Min(AllCount-PastCount,PageSize)
	%>
		<input type="hidden" name="page" value="<%=page%>"><%For i=1 To ShowCount
	If i mod 2 = 0 Then 
         bgcolor = "ArticleTr1"
       Else
         bgcolor = "ArticleTr2" 
     End If
     %>
		<tr onMouseOver="this.className='ahover'" onMouseOut="this.className='<%=bgcolor%>'" class="<%=bgcolor%>">
			<td width="33" align="center">&nbsp;&nbsp;&nbsp;&nbsp;<input name="artID" type="checkbox" onClick="unselectall()" id="artID" value="<%=rs("ID")%>"></td>
			<td width="680" align="left"><%
dim title,strlength
title=Rs("Title")
if forumid="" then strlength=40 else strlength=50
if strlen(title)>strlength then
title=CutStr(title,strlength)
end if
If Rs("TitleBold") = True Then title="<b>"&title&"</b>"
If Rs("TitleColor") <>"" Then
title="<font color="""&Rs("TitleColor")&""">"&title&"</font>"
Else
title=""&title&""
End If

'����Ŀǰ����ʾ��Ŀ���ƣ�������ʾ����Ŀ����
If forumID=""  Then    
Response.write "&nbsp;<a href=admin_show.asp?target=article&sqlstr=classid&forumID=" & Rs("classID") &  ">" & "<font color=#666666>[" & Application(systemkey&Rs("classID")) &"]</font>" & "</a>"& "&nbsp;"
Else
If Rs("subclass")>0 and Rs("classid")=Cnum(forumID) Then Response.write("<font color=#666666>["&subclassArray(Rs("subclass"))&"]</font>&nbsp;")
End If
Response.write "<a target=""_blank"" href=""../context.asp?id=" & Rs("ID") & """>"  &  title & "</a>"
If Rs("SetTop")=1 Then
	Response.write "&nbsp;&nbsp;<img align=absmiddle border=""0"" src=""images/thread_top.gif"" width=""15"" height=""15"" alt=""ר���ȵ�"">"
End If
If Rs("Classic")=1 Then
	Response.write "&nbsp;&nbsp;<img align=absmiddle border=""0"" src=""images/thread_digest.gif"" width=""15"" height=""15"" alt=""�����Ƽ�"">"
End If
If Rs("lock")=true Then
	Response.write "&nbsp;&nbsp;<img align=absmiddle border=""0"" src=""images/thread_swf.gif"" width=""15"" height=""15"" alt=""��ͼ�õ�Ƭ"">"
End If
			%></a></td>
			<td align="center" width="80">
			<a href="admin_all.asp?target=article&act=deletesingle&id=<%=rs("id")%>" onClick="javascript:return confirm('ȷ��ɾ��������?(��������һ�лظ�)');">
			ɾ��</a> - <a href="../apost.asp?act=edit&id=<%=rs("id")%>">�༭</a></td>
			<%rs.movenext%>
			</tr>
    <%Next
	%> 
		
		<tr class="a10">
			<td>&nbsp;&nbsp;&nbsp;&nbsp;<input name="chkAll" type="checkbox" id="chkAll" onClick="CheckAll(this.form)" value="checkbox"></td>
			<td colspan="3">&nbsp;&nbsp;<strong>������</strong>
			<input name="Action" type="radio" value="delete">ɾ�� |&nbsp;
			<input name="Action" type="radio" value="resetbold">��� |&nbsp;
			<input name="Action" type="radio" id="move" value="move" checked>�ƶ��� 
			<select onChange="javascript:setclass();" class="Sel" name="classid" id="classid"><%for i=1 to BoardCount
				Response.write "<option value='"&i&"'"
				If i=Cnum(forumID) Then
				Response.write " selected"
				End If
				Response.write ">"&Application(systemkey&i)&"</option>"
				Next
				Response.write "<option value='999'>"&Application(systemkey&999)&"</option>"
				rs.close%></select> <%
				Response.write("&nbsp;<select onchange=""javascript:document.all('move').checked=1;"" class=""Sel"" id=""subclass"" name=""subclass"">")
				For i=0 to ubound(subclassArray)
				Response.write("<option value="&i&">"&subclassArray(i)&"</option>")
				Next
				Response.write("</select>&nbsp;&nbsp;")
			%><input type="submit" name="Submit" value="ִ��" class="btn2">
			</td>
		</tr>
	</table>
</form>
</div>
<%
'��ʾ��������
Sql="select count(*) from content where parent=0"
Openrs(sql)
Response.Write("<p class=""ArticleUp""><span>����<font color=""red"">"&rs(0)&"</font>ƪ����&nbsp;&nbsp;</span>")

If page>1 Then
Response.Write("<a href=""admin_show.asp?target=article&forumid="&forumid&"&page="&page-1&"&subclassID="&subclassID&""">��һҳ</a>")
End If

Response.write ShowPages(Pages,Page,"admin_show.asp?target=article&forumid="&forumid&"&sqlstr="&sqlstr&"&subclassID="&subclassID&"")


If page<pages Then
Response.Write("<a href=""admin_show.asp?target=article&forumid="&forumid&"&page="&page+1&"&subclassID="&subclassID&""">��һҳ</a>")
End If
Response.Write("</p></div>")
'========================================================
'���۹���
	Case "Comments"
	forumid=Request.QueryString("forumid")
	sqlstr=Request.Querystring("sqlstr")
	subclassid=Request.QueryString("subclassid")
	'��ȡ����Ŀ
	tempClass=Application(SystemKey&"subclass"&ForumID)
	If instr(tempClass,"|")>0 Then
	subclassArray=split(tempClass,"|")
	Else
	subclassArray=split("����","|")
	End If
	if forumid="" then
		if sqlstr="" then    	
			orderStr=" order by settop desc,lastupdatetime desc"
    	end if
		whereStr=" where ReplyNum>0" 
	else
    	if sqlstr="classid" or sqlstr="" then
		orderStr= "  order by PostTime desc"
		end if		
		whereStr=" where ReplyNum>0 and classid="&forumid&"" 
    end if 
	If IsNumeric(subclassID)=True and Cnum(subclassID)>0 and Cnum(subclassID)<999 Then
	whereStr=whereStr&" and subclass="&subclassID&""
	End If
	If Cnum(subclassID)=999 Then
	whereStr=whereStr&" and subclass<1"
	End If
	sql="select * from content "&whereStr&orderStr
	openrs(sql)
		%>
<div class="DB">
    <div class="L1">��ǰλ�ã�<a href="admin_show.asp">�������</a> &gt; ���۹���</div>
	<div class="D3">
		 <div class="T2">
	<%If forumID="" Then%>
	<a href="admin_show.asp?target=Comments">ȫ������</a> <%For j=1 To BoardCount
				Print "| <a href=admin_show.asp?target=Comments&sqlstr=classid&forumID=" &j&  ">"&Application(systemkey&j)&"</a> "
			Next%>
	<%Else
	Response.write("&nbsp;<a href=admin_show.asp?target=Comments>ȫ������</a>&nbsp;|&nbsp;")
	for i=1 to ubound(subclassArray)
	If i=cnum(subclassID) Then 
	Response.write("&nbsp;-&nbsp;<b>"&subclassArray(i)&"</b>")
	Else
	response.write("&nbsp;-&nbsp;<a class='classnametitle' href=admin_show.asp?target=Comments&forumID="&forumID&"&subclassID="&i&">"&subclassArray(i)&"</a>")
	End If
	next
	response.write("&nbsp;-&nbsp;<a class='classnametitle' href=admin_show.asp?target=Comments&forumID="&forumID&"&subclassID=999>δ����</a>")
	End If%>
	</div>
	<table cellspacing="0" cellpadding="4" border="0" class="a23" bgcolor="#FFFFFF" align="center">
		<% 
	'�趨ÿҳ��ʾ������
	PageSize=18
    AllCount=Rs.RecordCount
	Pages=Fix(AllCount/PageSize)
	If Pages*PageSize<AllCount Then
		Pages=Pages+1	
	End If
	PastCount=(Page-1)*PageSize
	If PastCount>=AllCount Then
		Page=1
		PastCount=0
	End If
	If AllCount>PastCount Then
		Rs.Move PastCount
	End If
	ShowCount=Min(AllCount-PastCount,PageSize)
	%>
		<input type="hidden" name="page" value="<%=page%>"><%For i=1 To ShowCount
	If i mod 2 = 0 Then 
         bgcolor = "ArticleTr1"
       Else
         bgcolor = "ArticleTr2" 
     End If
     %>
		<tr onMouseOver="this.className='ahover'" onMouseOut="this.className='<%=bgcolor%>'" class="<%=bgcolor%>">
			<td width="440" align="left">&nbsp;&nbsp;<%
title=Rs("Title")
if forumid="" then strlength=40 else strlength=50
if strlen(title)>strlength then
title=CutStr(title,strlength)
end if
title=""&title&""

'����Ŀǰ����ʾ��Ŀ���ƣ�������ʾ����Ŀ����
If forumID=""  Then    
Response.write "&nbsp;<a href=admin_show.asp?target=Comments&sqlstr=classid&forumID=" & Rs("classID") &  ">" & "<font color=#666666>[" & Application(systemkey&Rs("classID")) &"]</font>" & "</a>"& "&nbsp;"
Else
If Rs("subclass")>0 and Rs("classid")=Cnum(forumID) Then Response.write("<font color=#666666>["&subclassArray(Rs("subclass"))&"]</font>&nbsp;")
End If

Response.write "<a target=""_blank"" href=""../context.asp?id=" & Rs("ID") & """>"&title&"</a>"
			%></a></td>
			<%
      rs.movenext
    Next%>
		</tr>
</table>
</form>
</div>
<%
Response.Write("<p class=""ArticleUp"">")
If page>1 Then
Response.Write("<a href=""admin_show.asp?target=Comments&forumid="&forumid&"&page="&page-1&"&subclassID="&subclassID&""">��һҳ</a>")
End If
Response.write ShowPages(Pages,Page,"admin_show.asp?target=Comments&forumid="&forumid&"&sqlstr="&sqlstr&"&subclassID="&subclassID&"")
If page<pages Then
Response.write("<a href=""admin_show.asp?target=Comments&forumid="&forumid&"&page="&page+1&"&subclassID="&subclassID&""">��һҳ</a>")
End If
Response.write("</p></div>")
'========================================================
'���ӹ���
Case "friendsite"
%>
<form name="FrontPage_Form9" method="POST" action="admin_all.asp?target=friendsite&act=new" onSubmit="return FrontPage_Form9_Validator(this)" language="JavaScript">
<div class="DB">
    <div class="L1">��ǰλ�ã�<a href="admin_show.asp">�������</a> &gt; �������ӹ���</div>
	<table cellspacing="1" cellpadding="5" border="0" class="a5" id="table1">
		<tr class="T4">
			<td width="29" class="center bold">ID</td>
			<td width="160" class="center bold">վ������</td>
			<td width="281" class="center bold">վ���ַ</td>
			<td width="116" class="center bold">����</td>
		</tr>
		<%sql="select * from friendsite order by siteorder"
    openrs(sql)
	For i=1 To Rs.RecordCount
	If i mod 2 = 0 Then 
          bgcolor = "Tr1"
        Else
         bgcolor = "Tr2" 
     End If
     %>
		<tr onMouseOver="this.className='ahover'" onMouseOut="this.className='<%=bgcolor%>'" class="<%=bgcolor%>">
			<td width="29" class="center"><%=rs("siteorder")%></td>
			<td width="160" class="center"><%=rs("sitename")%></td>
			<td width="281" class="center">
			<a href="<%=rs("siteurl")%>" target="_blank"><%=rs("siteurl")%></a></td>
			<td class="center" width="116">
			<p>
			<a href="admin_all.asp?target=friendsite&act=edit&siteorder=<%=rs("siteorder")%>">
			�޸�</a> <span lang="en-us">&nbsp; </span>
			<a href="admin_all.asp?target=friendsite&act=delete&siteorder=<%=rs("siteorder")%>" onClick="javascript:return askdelete();">
			ɾ��</a></p>
			</td>
		</tr>
		<%
	rs.movenext
	Next
	%>
		<tr class="a10">
			<td width="29" class="center"><input type="text" name="siteorder" size="3" maxlength="4" class="IT3"></td>
			<td width="160" class="center"><input type="text" name="sitename" size="21" maxlength="30" class="IT2"></td>
			<td width="281" class="center"><input type="text" name="siteurl" size="38" maxlength="100" class="IT5" value="http://"></td>
			<td class="center"><input type="submit" value="ȷ�����" name="addurl" class="btn3"></td>
		</tr>
	</table></div>
</form>
<%
'========================================================
'ҳͷͼƬ����
Case "topbg"
%>
<div class="DB">
  <div class="L1">��ǰλ�ã�<a href="admin_show.asp">�������</a> &gt; ҳͷͼƬ����</div>
	<table cellspacing="1" cellpadding="5" border="0" class="a5" id="table1">
		<tr>
			<td width="120" class="center bold">ͼƬԤ��</td>
			<td width="500"><img src="../images/topbg.jpg" width="500"></td>
		</tr>
        <tr>
			<td colspan="2">ע�⣺<font color="#FF0000"> * ͼƬ������Ϊ��topbg.jpg�������ϴ�������,���ϴ��Ľ����Ǿɵġ�ͼƬ��С1250px * 680px��<a href="../images/topbg.psd">�������psdģ��</a>)</font></td>
		</tr>
        <tr> 
		<td colspan="2" align="center"><iframe name="ad" frameborder=0 width=100% height=60 scrolling=auto src=uploadc.asp></iframe></td>
	</tr>
	</table></div>

<%
'=========================================================
'�鿴������־
Case "SetLog"
%>
<div class="DB">
<div class="L1">��ǰλ�ã�<a href="admin_show.asp">�������</a> &gt; ������־</div>
<table cellspacing="1" cellpadding="3" bgcolor="#cccccc" width="650">
<tr class="T4">
<td width="100" class="center bold">����Ա����</td>
<td width="400" class="center bold">��������</td>
<td width="150" class="center bold">����ʱ��</td>
</tr>
<%
Dim AdminName
sqlstr=Request.Querystring("sqlstr")
AdminName=Request.Querystring("AdminName")
Select Case sqlstr
case "name" 
sql="select * from SetLog where AdminName='"&AdminName&" order by id desc'"
case else
sql="select * from SetLog order by id desc"
End Select
Openrs(sql)

PageSize=18
AllCount=Rs.RecordCount
Pages=Fix(AllCount/PageSize)
If Pages*PageSize<AllCount Then
Pages=Pages+1	
End If
PastCount=(Page-1)*PageSize
If PastCount>=AllCount Then
Page=1
PastCount=0
End If
If AllCount>PastCount Then
Rs.Move PastCount
End If
ShowCount=Min(AllCount-PastCount,PageSize)

For i=1 To ShowCount
If i mod 2 = 0 Then 
bgcolor = "Tr1"
Else
bgcolor = "Tr2" 
End If
%>
<tr onMouseOver="this.className='ahover'" onMouseOut="this.className='<%=bgcolor%>'" class="<%=bgcolor%>">
<td width="100" class="center"><a href="?target=SetLog&sqlstr=name&AdminName=<%=rs("AdminName")%>"><%=rs("AdminName")%></a></td>
<td width="400">&nbsp;&nbsp;<%=rs("SetTxt")%></td>
<td width="150" class="center"><%=rs("SetTime")%></td>
</tr>
<%
Rs.Movenext
Next
Rs.Close
%>
</table>

<p class="Pages2">
<span>��ʾ���������Ա���ƿ��Բ鿴�ù���Ա�Ĳ�����־</span>
<%If page>1 Then%>
<a href="?target=SetLog&page=<%=page-1%>&sqlstr=<%=sqlstr%>&AdminName=<%=AdminName%>">��һҳ</a>&nbsp;
<%End If%>
<%
Response.write ShowPages(Pages,Page,"?target=SetLog&sqlstr="&sqlstr&"&AdminName="&AdminName&"")
%>
<%If page<pages Then%>
&nbsp;<a href="?target=SetLog&page=<%=page+1%>&sqlstr=<%=sqlstr%>&AdminName=<%=AdminName%>">��һҳ</a>
<%End If%>
</p>
</div>
<%
'=========================================================
'Ͷ�����
Case "Submission"
%>
<div class="DB">
<div class="L1">��ǰλ�ã�<a href="admin_show.asp">�������</a> &gt; Ͷ�����</div>
<table cellspacing="1" cellpadding="4" bgcolor="#cccccc" width="650">
<tr class="T4">
<td width="100" class="center bold">Ͷ���Ա</td>
<td width="300" class="center bold">���±���</td>
<td width="150" class="center bold">Ͷ��ʱ��</td>
<td width="100" class="center bold">����</td>
</tr>
<%
sqlstr=Request.Querystring("sqlstr")
AdminName=Request.Querystring("AdminName")
Select Case sqlstr
case "name" 
sql="select * from Content"
case else
sql="select * from Content where Submission=-1 order by Verify asc,id desc"
End Select
Openrs(sql)

PageSize=18
AllCount=Rs.RecordCount
Pages=Fix(AllCount/PageSize)
If Pages*PageSize<AllCount Then
Pages=Pages+1	
End If
PastCount=(Page-1)*PageSize
If PastCount>=AllCount Then
Page=1
PastCount=0
End If
If AllCount>PastCount Then
Rs.Move PastCount
End If
ShowCount=Min(AllCount-PastCount,PageSize)

For i=1 To ShowCount
If i mod 2 = 0 Then 
bgcolor = "Tr1"
Else
bgcolor = "Tr2" 
End If
%>
<tr onMouseOver="this.className='ahover'" onMouseOut="this.className='<%=bgcolor%>'" class="<%=bgcolor%>">
<td width="100" class="center"><%=rs("Poster")%></td>
<td width="300">&nbsp;&nbsp;<a href="../context.asp?id=<%=rs("id")%>" target="_blank"><%=rs("Title")%></a></td>
<td width="150" class="center"><%=rs("PostTime")%></td>
<td width="100" class="center">
<%If Rs("Verify")=-1 Then%>
<a href="admin_all.asp?target=article&act=deletesingle&id=<%=rs("id")%>" onClick="javascript:return confirm('ȷ��ɾ��������?(��������һ�лظ�)');">ɾ��</a> - <a href="../apost.asp?act=edit&id=<%=rs("id")%>">�༭</a>
<%Else%>
<font color="#cccccc">�����</font>
<%End If%>
</td></tr>
<%
Rs.Movenext
Next
Rs.Close
%>
</table>

<p class="Pages2">
<span>��ʾ���Ѿ�ͨ����˵����£��뵽���¹�����й���</span>
<%If page>1 Then%>
<a href="?target=Submission&page=<%=page-1%>">��һҳ</a>&nbsp;
<%End If%>
<%
Response.write ShowPages(Pages,Page,"?target=Submission")
%>
<%If page<pages Then%>
&nbsp;<a href="?target=Submission&page=<%=page+1%>">��һҳ</a>
<%End If%>
</p>
</div>
<%
'=========================================================
'������վ��ҳ
Case "ShengCheng"
dim objXmlHttp,binFileData,objAdoStream 
set objXmlHttp = Server.CreateObject("Microsoft.XMLHTTP") 
objXmlHttp.open "GET","http://"&request.ServerVariables("HTTP_HOST")&"/index.asp",false 
objXmlHttp.send() 
binFileData = objXmlHttp.responseBody 
set objAdoStream = Server.CreateObject("ADODB.Stream") 
objAdoStream.Type = 1 
objAdoStream.Open() 
objAdoStream.Write(binFileData) 
objAdoStream.SaveToFile server.MapPath("../index.html"),2 
objAdoStream.Close() 
set objAdoStream=nothing 
set objXmlHttp=nothing 
'��������Ա������־
Sql="select * from SetLog"
OpenRs(Sql)
Rs.AddNew
Rs("AdminName")=User_ID
Rs("SetTxt")="������վ��ҳ"
Rs("SetTime")=Now
Rs.Update
Rs.Close
'������ɷ���
Response.write"<script language=javascript>alert('��ҳ�������!');this.location.href='admin_show.asp';</script>"

Case Else

End Select 
Call CloseAll()
%>
</body>
</html>
