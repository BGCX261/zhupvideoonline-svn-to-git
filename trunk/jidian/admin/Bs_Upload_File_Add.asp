<!--#include file="bsconfig.asp"-->
<!-- #include file="Inc/Head.asp" -->
<%call checkmanage("64")%>
<body onLoad="javascipt:setTimeout('loadForm()',1000);">
<table align="center" class="a2">
	<tr>
		<td class="a1">添加网络课程</td>
	</tr>
	<tr class="a4">
		<td>
						<table width="100%">
							<tr> 
								<td class="iptlab">说明：</td>
								<td class="ipt">此权限慎用！主要是用来向网站传送大文件。用于在空间保存，固定在网站的UploadFilesDown目录下</td>
							</tr>
                            <tr>
							  <td class="iptlab">大文件上传[FLV，rar，zip格式]:</td>
							  <td class="ipt"><iframe class="TBGen" style="top:2px" ID="Uploaddown" src="../Progress_up.asp" frameborder=0 scrolling=no width="380" height="60"></iframe></td>
							</tr>
						</table>

		</td>
	</tr>
</table>
<%htmlend%>
