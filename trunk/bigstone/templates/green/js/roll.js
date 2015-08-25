var roll_width = 520;
var roll_height = 178;
var unit_width = 150;
var unit_height = 178;
var unit_margin = 7;
var roll_px = 1;
var roll_interval = 30;
var roll_i = 1;
var roll_j = 1;
var sheet_width = 0;
document.getElementById("roll").style.width = roll_width + "px";
document.getElementById("roll").style.height = roll_height + "px";
code = document.getElementById("roll_sheet").innerHTML;
for(i = 0;i < 3;i ++)
{
	code += code;
}
document.getElementById("roll_sheet").innerHTML = code;
if(navigator.appName == "Microsoft Internet Explorer")
{
  unit_count = document.getElementById("roll_sheet").childNodes.length;
  for(i = 0;i < unit_count;i ++)
  {
	obj = document.getElementById("roll_sheet").childNodes.item(i).style;
	obj.width = unit_width + "px";
	obj.height = unit_height + "px";
	obj.marginLeft = unit_margin + "px";
	obj.marginRight = unit_margin + "px";
  }
  sheet_width = (unit_width + unit_margin * 2) * unit_count + unit_margin;
  document.getElementById("roll_sheet").style.width = sheet_width + "px";
}else{
  unit_count = document.getElementsByName("roll_unit").length;
  for(i = 0;i < unit_count;i ++)
  {
	obj = document.getElementsByName("roll_unit").item(i).style;
	obj.width = unit_width + "px";
	obj.height = unit_height + "px";
	obj.marginLeft = unit_margin + "px";
	obj.marginRight = unit_margin + "px";
  }
  sheet_width = (unit_width + unit_margin * 2) * unit_count + unit_margin;
  document.getElementById("roll_sheet").style.width = sheet_width + "px";
}
function roll()
{
  if(roll_i < sheet_width - roll_width)
  {
	document.getElementById("roll_sheet").style.left = -roll_i + "px";
	roll_j = (roll_i = roll_i + roll_px);
  }
  else
  {
	document.getElementById("roll_sheet").style.left = -roll_j + "px";
	roll_j = roll_j - roll_px;
	if(roll_j < 0) roll_i = roll_j;
  }
}
var setroll = setInterval("roll()",roll_interval);
function over_roll()
{
  clearInterval(setroll);
}
function out_roll()
{
  setroll = setInterval("roll()",roll_interval);
}
//新秀