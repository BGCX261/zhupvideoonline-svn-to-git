<script language=javascript>
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
</script> 
<div class="leftbox">
    <div class="leftboxtitle">ÍøÂç¿Î³Ì</div>
    <div class="leftboxcenter"><% call ShowAllClassl() %></div>
    <div class="leftboxbottom"></div>
</div>
