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
      <td class="bodyleft"><div id="leftbg">
          <!-- #include file="L_product.asp" -->
        </div></td>
      <td class="bodyline"></td>
      <td class="bodyright"><table class="mbox">
          <tr>
            <td><div class="rtitle"><img src="img/title_ico.gif" />视 频 作 品 搜 索</div></td>
          </tr>
          <tr>
            <td><table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td height="28"><% call ShowSearchTerm() %>
                  </td>
                  <td width="100"><% call SearchResultTotal() %>
                  </td>
                </tr>
              </table></td>
          </tr>
          <tr>
            <td class="cttd"><div class="content">
                <% call ShowSearchResult() %>
            </div></td>
          </tr>
          <tr>
            <td id='tpage'><%
if totalput>0 then
call showpage(strFileName,totalput,MaxPerPage,false,true,"个视频")
end if
%>
            </td>
          </tr>
        </table></td>
    </tr>
  </table>
</div>
<!-- #include file="Inc/Foot.asp" -->
</body>
</html>
