<!--#include file="Inc/syscode.asp"-->
<!--#include file="Inc/Top.asp" -->
<%
'请勿改动下面这三行代码
ShowSmallClassType=ShowSmallClassType_Default
MaxPerPage=MaxPerPage_Default
strFileName="article.asp?BigClassName=" & Server.URLEncode(BigClassName) & "&SmallClassName=" & Server.URLEncode(SmallClassName) 
%>
<!-- #include file="Inc/Head.asp" -->
<div id="body">
  <table class="bodytable">
    <tr>
      <td class="bodyleft">
          <!-- #include file="L_article.asp" -->
      </td>
      <td class="bodyline"></td>
      <td class="bodyright">
      <!-- #include file="Inc/headAd.asp" -->
        <div class="mbox">
            <div class="artlist">
            <div class="artpath"><span style=" float:left;"><img src="img/title_ico.gif" /> <% call ShowClassGuide() %></span> <span style=" float:right;"><% call ShowArticleTotal() %></span> </div>
			<% call ShowArticle(32) %>
            </div>
            <div id='tpage'><%
if totalput>0 then
call showpage(strFileName,totalput,MaxPerPage,false,true,"条记录")
end if
%>            
            </div>
        </div>
      </td>
    </tr>
  </table>
</div>
<!-- #include file="Inc/Foot.asp" -->
</body>
</html>
