<!--#include file="Inc/syscode.asp"-->
<%
ShowSmallClassType=ShowSmallClassType_Search
MaxPerPage=MaxPerPage_Search
strFileName="Sch.asp?Field=" & strField & "&Keywd=" & keywd & "&nbtyped=" & nbtyped
%>
<!-- #include file="Inc/Head.asp" -->
<div id="body">
  <table class="bodytable">
    <tr>
      <td class="bodyleft">
          <!-- #include file="L_stuZone.asp" -->
      </td>
      <td class="bodyline"></td>
      <td class="bodyright">
      <!-- #include file="Inc/headAd.asp" -->
        <div class="mbox">
            <div class="positiontitle"><img src="img/title_ico.gif" /> ʦ�������������</div>
            <div class="artlist">
              <div class="artpath"><span style=" float:left;"><% call ShowSchTerm() %></span> <span style=" float:right;"><% call SchResultTotal() %></span> </div>
            </div>
            <div class="content">
                <% call ShowSchResult() %>
            </div>
            <div id='tpage'>
<%
if tolput>0 then
call showpage(strFileName,tolput,MaxPerPage,false,true,"����Ϣ")
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
