<SCRIPT language=javascript1.2>
function showsubmenu(sid)
{
whichEl = eval("submenu" + sid);
if (whichEl.style.display == "none")
{
eval("submenu" + sid + ".style.display=\"\";");
}
else
{
eval("submenu" + sid + ".style.display=\"none\";");
}
}
</SCRIPT> 
<table class="leftbox">
  <tr>
    <td class="leftboxtitle">视频作品分类</td>
  </tr>
  <tr>
    <td class="leftboxcenter"><% call ShowAllClassl() %></td>
  </tr>
  <tr>
    <td class="leftboxbottom"></td>
  </tr>
</table>
