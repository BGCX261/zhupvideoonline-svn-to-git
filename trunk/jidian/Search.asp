<!--#include file="Inc/syscode.asp"-->
<%
ShowSmallClassType=ShowSmallClassType_Search
MaxPerPage=MaxPerPage_Search
strFileName="Search.asp?Field=" & strField & "&Keyword=" & keyword & "&BigClassName=" & BigClassName & "&SmallClassName=" & SmallClassName 
%>
<!-- #include file="Inc/Head.asp" -->
<%if not(rsBigClass.bof and rsBigClass.eof) and ShowSmallClassType="Menu" then response.write " onmousemove='HideMenu()'"%>
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
            <div class="positiontitle"><img src="img/title_ico.gif" />я╖о╟ндубкякВ</div>
            <div class="artlist">
              <div class="artpath"><span style=" float:left;"><% call ShowSearchTerm() %></span> <span style=" float:right;"><% call SearchResultTotal() %></span> </div>
            </div>
            <div class="content">
                <% call ShowSearchResult() %>
            </div>
            <div id='tpage'><%
if totalput>0 then
call showpage(strFileName,totalput,MaxPerPage,false,true,"ф╙ндуб")
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
