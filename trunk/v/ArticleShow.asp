<!--#include file="Inc/syscode.asp"-->
<!--#include file="Inc/top.asp"-->
<%
ShowSmallClassType=ShowSmallClassType_Article
dim ArticleID
ArticleID=trim(request("ArticleID"))
if ArticleId="" then
	response.Redirect("Product.asp")
end if

sql="select * from Product where ArticleID=" & ArticleID & ""
Set rs= Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,3
if rs.bof and rs.eof then
	response.write"<SCRIPT language=JavaScript>alert('找不到此产品！');"
  response.write"javascript:history.go(-1)</SCRIPT>"
else	
	rs("Hits")=rs("Hits")+1
	rs.update
	if rs("hits")>=HitsOfHot then
		rs("Hot")=True
		rs.update
	end if
	BigClassName=rs("BigClassName")
	SmallClassName=rs("SmallClassName")
%>
<%
dim selclass
selclass=rs("BigClassName")
sqlClassWrd="select * from BigClass where BigClassName='" & selclass & "'"
Set rsClassWrd= Server.CreateObject("ADODB.Recordset")
rsClassWrd.open sqlClassWrd,conn,1,1
%>
<!-- #include file="Inc/Head.asp" -->
<div id="body">
  <table class="bodytable">
    <tr>
      <td class="bodyleft"><div id="leftbg">
          <!-- #include file="L_product.asp" -->
        </div></td>
      <td class="bodyline"></td>
      <td class="bodyright"><table class="mbox">
          <tr>
            <td><div class="rtitle"><img src="img/title_ico.gif" /> 视 频 作 品 展 示</div></td>
          </tr>
          <tr>
            <td><%
response.write "|&nbsp;<a href='Product.asp?BigClassName=" & rs("BigClassName") & "'>" & rs("BigClassName") & "</a>&nbsp;&gt;&gt;&nbsp;"
if rs("SmallClassName") & ""<>"" then
response.write "<a href='Product.asp?BigClassName=" & rs("BigClassName")&"&SmallClassName=" & rs("SmallClassName") & "'>" & rs("SmallClassName") & "</a>&nbsp;&gt;&gt;&nbsp;"
end if
response.write rs("Title")
%>
            </td>
          </tr>
          <tr>
            <td><table class="mytable">
                <tr>
                  <td  class="chptitletd"><%=rs("Title")%></td>
                </tr>
                <tr>
                  <td  class="chphittd">访问数：<%=rs("Hits")%>&nbsp;评论(人气)数：<%=rs("Votes")%>&nbsp; 录入时间：<%= FormatDateTime(rs("UpdateTime"),2) %></td>
                </tr>
<%
				if rsClassWrd("Wrd1")<>"" then
%>
                <tr>
                  <td class="chpparttd"><span><%=trim(rsClassWrd("Wrd1"))%></span>
                  <%=rs("Zi1")%></td>
                </tr>
<%
end if
if rsClassWrd("Wrd2")<>"" then	
%>	
                <tr>
                  <td class="chpparttd"><span><%=trim(rsClassWrd("Wrd2"))%></span>
                  <%=rs("Zi2")%></td>
                </tr>
<%
end if
if rsClassWrd("Wrd3")<>"" then	
%>	
                <tr>
                  <td class="chpparttd"><span><%=trim(rsClassWrd("Wrd3"))%></span>
                  <%=rs("Zi3")%></td>
                </tr>
				<%
end if
if rsClassWrd("Wrd4")<>"" then	
%>	
                <tr>
                  <td class="chpparttd"><span><%=trim(rsClassWrd("Wrd4"))%></span>
                  <%=rs("Zi4")%></td>
                </tr>
				<%
end if
if rsClassWrd("Wrd5")<>"" then	
%>	
                <tr>
                  <td class="chpparttd"><span><%=trim(rsClassWrd("Wrd5"))%></span>
                  <%=rs("Zi5")%></td>
                </tr>
				<%
end if
if rsClassWrd("Wrd6")<>"" then	
%>	
                <tr>
                  <td class="chpparttd"><span><%=trim(rsClassWrd("Wrd6"))%></span>
                  <%=rs("Zi6")%></td>
                </tr>
				<%
end if
if rsClassWrd("Wrd7")<>"" then	
%>	
                <tr>
                  <td class="chpparttd"><span><%=trim(rsClassWrd("Wrd7"))%></span>
                  <%=rs("Zi7")%></td>
                </tr>
				
<%
end if
rsClassWrd.close
%>
                <tr>
                  <td class="chppartVtd" >

                    <%if rs("UploaddownJ")<>"" then%>
  <script type="text/javascript" src="js/swfobject.js"></script>
  <p id="player1"><a href="http://www.macromedia.com/go/getflashplayer">下载Flash播放器</a>观看flash动画.</p>
  <script type="text/javascript">
	var s1 = new SWFObject("flvplayer.swf","single","<%=rs("flvWidth")%>","<%=rs("flvHight")+20%>","7");
	s1.addParam("allowfullscreen","true");
	s1.addVariable("file","<%=rs("UploaddownJ")%>");
	s1.addVariable("image","<%=rs("Uploaddown")%>");
	s1.addVariable("width","<%=rs("flvWidth")%>");
	s1.addVariable("height","<%=rs("flvHight")+20%>");
	s1.write("player1");
</script>
                    <%=rs("downsizeJ") %> Kb
                  <%end if%>	  </td>
                </tr>
				<tr>
                  <td><div class="chpviewtd"><%call ShowArticleContent()%>                  </div></td>
                </tr>
                <tr>
                  <td>
                  <%
randomize
yzm=int(8999*rnd()+1000)

set rs=conn.execute("select * from main")
if rs("Review")=True then '是否开放评论
%>

<script language="javascript">
<!--
function check()
{
	if (document.pinglunform.name.value.length < 3 || document.pinglunform.name.value.length >16) {
		alert("出错了，请输入有效的用户名3-16位之间。");
		document.pinglunform.name.focus();
		return false;
	}
    if (document.pinglunform.yzm.value.length <1){
		alert("出错了，您没有填写验证码。");
		document.pinglunform.yzm.focus();
		return false;
	}
    <%
response.write "if (document.pinglunform.yzm.value!="&yzm&"){"
%>
		alert("出错了，您填写的验证码不正确。");
		document.pinglunform.yzm.focus();
		return false;
	}
	if (document.pinglunform.nr.value.length < 5) {
		alert("出错了，请您填写评论内容不得少于5字。");
		document.pinglunform.nr.focus();
		return false;
	}
    if (document.pinglunform.nr.value.length >200) {
		alert("出错了，您的评论内容太长了。");
		document.pinglunform.nr.focus();
		return false;
	}
	
	var res = confirm("是否真的要提交？");
        if(res == true){
	    document.pinglunform.target="_self";
		document.pinglunform.action="pinglun_save.asp";
	    document.pinglunform.submit();
        }

}
//-->
</script>
<%


  set rs=conn.execute("select top 5 * from buyok_pinglun where prodid='"&ArticleID&"' order by AddDate desc")
  if rs.eof and rs.bof then
  response.write "<div>暂时没有用户对此作品提交评论。</div>"
  else
  do while not rs.eof
%>
     
<%
if rs("shenhe")=True then
   response.write "<table class='pingLunBox'><tr class='pingLunTitle'><td width=33% title="&(rs("IP"))&">作者："&rs("name")&"</td><td>发表时间："&rs("adddate")&"</td><td>"
   for j=1 to rs("jibie")
   response.write "<font color=orange>☆</font>"
   next
  response.write "</td></tr>" 
  response.write "<tr><td colspan=3><img src=img/mood/" &rs("face")&" /> <span style='line-height=150%'>"&replace(server.htmlencode(rs("nr")),vbCRLF,"<BR>")&"</span></td></tr></table>"
    
else

  response.write "<table  class='pingLunBox'><tr class='pingLunTitle'><td width=33% title="&(rs("IP"))&"> 作者："&rs("name")&"</td><td>发表时间："&rs("adddate")&"</td></tr>"
  response.write "<tr><td colspan=3>信息已提交，管理员审核后才能显示</td></tr></table>"
end if 

  rs.movenext
  loop
  response.write "<div><a href=pinglun.asp?ArticleID="&ArticleID&" target='_blank'>浏览该作品的全部评论>></a></div>"
  end if
%>


</td></tr>

<tr><td align=center>
	<form name="pinglunform" method="post">
	<table style="border:#ccc solid 1px; width:100%;">
		<tr>
		<td width="8%">姓 名：</td>
		<td><input name="name" id="name" size="16" maxLength="16" readonly="readonly" <%if session("UserName") <> "" then response.write "value='" & session("UserName") &"'" else response.write "value='需登录才能发表评论' onclick=javascript:window.location='Cn_index.asp'" end if %>>
        验证码：<input id="yzm" size="4" maxlength=4 name="yzm">
<%
a=int(yzm/1000)
b=int((yzm-a*1000)/100)
c=int((yzm-a*1000-b*100)/10)
d=int(yzm-a*1000-b*100-c*10)
response.write "<img src=img/yzm/"&a&".gif><img src=img/yzm/"&b&".gif><img src=img/yzm/"&c&".gif><img src=img/yzm/"&d&".gif>"
%>
        <font color=orange> 
		<input type="radio" value="1" name="jibie">☆
		<input type="radio" value="2" name="jibie">☆☆
		<input type="radio" value="3" checked name="jibie">☆☆☆
		<input type="radio" value="4" name="jibie">☆☆☆☆
		<input type="radio" value="5" name="jibie">☆☆☆☆☆</font></td>
		</tr>
        <tr>
          <td>表 情：</td>
          <td>
                            <input type="radio" value="mood1.gif" name="face" checked="checked" />
                            <img src="img/mood/mood1.gif" />
                            <input type="radio" value="mood2.gif" name="face" />
                            <img src="img/mood/mood2.gif" />
                            <input type="radio" value="mood3.gif" name="face" />
                            <img src="img/mood/mood3.gif"  />
                            <input type="radio" value="mood4.gif" name="face" />
                            <img src="img/mood/mood4.gif"  />
                            <input type="radio" value="mood5.gif" name="face" />
                            <img src="img/mood/mood5.gif"  />
                            <input type="radio" value="mood6.gif" name="face" />
                            <img src="img/mood/mood6.gif" />
                            <input type="radio" value="mood7.gif" name="face" />
                            <img src="img/mood/mood7.gif" />
                            <input type="radio" value="mood8.gif" name="face" />
                            <img src="img/mood/mood8.gif" />
                            <input type="radio" value="mood12.gif" name="face" />
                            <img src="img/mood/mood12.gif"/>
                            <input type="radio" value="mood9.gif" name="face" />
                            <img src="img/mood/mood9.gif" />
                            <input type="radio" value="mood11.gif" name="face" />
                            <img src="img/mood/mood11.gif"  />
                            <input type="radio" value="mood10.gif" name="face" />
                            <img src="img/mood/mood10.gif" />
		  </td>
        </tr>
		<tr>
		<td vAlign="top">评 论：<br><font color=red>≤200字</font></td>
		<td><textarea id="nr" name="nr" rows="4" cols="70"></textarea></td>
        </tr>
		<tr>
		<td></td>
		<td>
		<input type=hidden name='prodid' value='<%=ArticleID%>'>
		<input type=hidden name='save' value='ok'>
		<input type="button" value=" 提交 " name="submit1" <%if session("UserName") <> "" then response.write "onClick=javascript:check();" else response.write "onclick=javascript:window.location='Cn_index.asp'" end if %>>  
		<input type="reset" value=" 重写 " name="Submit2">
		</td></tr>
	</table>
    </form>
<%
else
  response.write "评论功能已被关闭" 
end if%>
                  
                  </td>
                </tr>
              </table></td>
          </tr>
          <tr>
            <td class="prt">【<a href='javascript:history.back()'>返回</a>】 </td>
          </tr>
        </table></td>
    </tr>
  </table>
</div>
<!-- #include file="Inc/Foot.asp" -->
</body>
</html>
<%
rsClassWrd.close
set rsClassWrd=nothing
%>
<%
end if
rs.close
set rs=nothing
call CloseConn()
%>
