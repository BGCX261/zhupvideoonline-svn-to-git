
<html>
<head>
<title>HTML���߱༭��</title>
<link rel="STYLESHEET" type="text/css" href="edit.css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body bgcolor="menu" STYLE="margin:0pt;padding:0pt">
<div class="yToolbar"> 
  <div class="TBHandle"> </div>
  <div class="Btn" TITLE="ɾ��" LANGUAGE="javascript" onClick="format('delete')"> 
    <img class="Ico" src="images\delete.gif" WIDTH="16" HEIGHT="16"> </div>
  <div class="TBSep"></div>
  <div class="Btn" TITLE="����" LANGUAGE="javascript" onClick="format('copy')"> 
    <img class="Ico" src="images\copy.gif" WIDTH="16" HEIGHT="16"> </div>
  <div class="Btn" TITLE="����" LANGUAGE="javascript" onClick="format('cut')"> <img class="Ico" src="images\cut.gif" WIDTH="16" HEIGHT="16"> 
  </div>
  <div class="Btn" TITLE="ճ��" LANGUAGE="javascript" onClick="format('paste')"> 
    <img class="Ico" src="images\paste.gif" WIDTH="16" HEIGHT="16"> </div>
  <div class="TBSep"></div>
  <div class="Btn" TITLE="����" LANGUAGE="javascript" onClick="format('undo')"> 
    <img class="Ico" src="images\undo.gif" WIDTH="16" HEIGHT="16"> </div>
  <div class="Btn" TITLE="�ָ�" LANGUAGE="javascript" onClick="format('redo')"> 
    <img class="Ico" src="images\redo.gif" WIDTH="16" HEIGHT="16"> </div>
  <div class="TBSep"></div>
  <div class="Btn" TITLE="�������" LANGUAGE="javascript" onClick="InsertTable()"> 
    <img class="Ico" src="images\table.gif" WIDTH="16" HEIGHT="16"> </div>
  <div class="Btn" TITLE="���볬������" LANGUAGE="javascript" onClick="UserDialog('CreateLink')"> 
    <img class="Ico" src="images\wlink.gif" WIDTH="16" HEIGHT="16"> </div>
  <div class="Btn" TITLE="����ˮƽ��" LANGUAGE="javascript" onClick="format('InsertHorizontalRule')"> 
    <img class="Ico" src="images\hr.gif" WIDTH="16" HEIGHT="16"> </div>
  <div class="Btn" TITLE="����ͼƬURL" LANGUAGE="javascript" onClick="InsertImg()"> 
    <img class="Ico" src="images\img.gif" WIDTH="16" HEIGHT="16"> </div>
  <div class="TBSep"></div>
  <iframe class="TBGen" style="top:2px" ID="UploadFiles" src="uploada.asp" frameborder=0 scrolling=no width="250" height="25"></iframe>
</div>

<div class="yToolbar"> 
  <div class="TBHandle"> </div>
  <select ID="formatSelect" class="TBGen" onChange="format('FormatBlock',this[this.selectedIndex].value);this.selectedIndex=0">
    <option selected>�����ʽ</option>
    <option VALUE="&lt;P&gt;">��ͨ</option>
    <option VALUE="&lt;PRE&gt;">�ѱ��Ÿ�ʽ</option>
    <option VALUE="&lt;H1&gt;">����һ</option>
    <option VALUE="&lt;H2&gt;">�����</option>
    <option VALUE="&lt;H3&gt;">������</option>
    <option VALUE="&lt;H4&gt;">������</option>
    <option VALUE="&lt;H5&gt;">������</option>
    <option VALUE="&lt;H6&gt;">������</option>
    <option VALUE="&lt;H7&gt;">������</option>
  </select>
  <select id="specialtype" class="TBGen" onChange="specialtype(this[this.selectedIndex].value);this.selectedIndex=0">
    <option selected>�����ʽ</option>
    <option VALUE="SUP">�ϱ�</option>
    <option VALUE="SUB">�±�</option>
    <option VALUE="DEL">ɾ����</option>
    <option VALUE="BLINK">��˸</option>
    <option VALUE="BIG">��������</option>
    <option VALUE="SMALL">��С����</option>
  </select>
  <div class="TBSep"></div>
  <div class="Btn" TITLE="�����" NAME="Justify" LANGUAGE="javascript" onClick="format('justifyleft')"> 
    <img class="Ico" src="images\aleft.gif" WIDTH="16" HEIGHT="16"> </div>
  <div class="Btn" TITLE="����" NAME="Justify" LANGUAGE="javascript" onClick="format('justifycenter')"> 
    <img class="Ico" src="images\center.gif" WIDTH="16" HEIGHT="16"> </div>
  <div class="Btn" TITLE="�Ҷ���" NAME="Justify" LANGUAGE="javascript" onClick="format('justifyright')"> 
    <img class="Ico" src="images\aright.gif" WIDTH="16" HEIGHT="16"> </div>
  <div class="TBSep"></div>
  <div class="Btn" TITLE="���" LANGUAGE="javascript" onClick="format('insertorderedlist')"> 
    <img class="Ico" src="images\numlist.gif" WIDTH="16" HEIGHT="16"> </div>
  <div class="Btn" TITLE="��Ŀ����" LANGUAGE="javascript" onClick="format('insertunorderedlist')"> 
    <img class="Ico" src="images\bullist.gif" WIDTH="16" HEIGHT="16"> </div>
  <div class="Btn" TITLE="����������" LANGUAGE="javascript" onClick="format('outdent')"> 
    <img class="Ico" src="images\outdent.gif" WIDTH="16" HEIGHT="16"> </div>
  <div class="Btn" TITLE="����������" LANGUAGE="javascript" onClick="format('indent')"> 
    <img class="Ico" src="images\indent.gif" WIDTH="16" HEIGHT="16"> </div>
  <div class="TBSep"></div>
  <div class="Btn" TITLE="�鿴����" LANGUAGE="javascript" onClick="help()"> <img class="Ico" src="images\help.gif" WIDTH="16" HEIGHT="16"> 
  </div>
  <!--<div class="TBSep"></div>
  <div class="Btn" TITLE="����"	LANGUAGE="javascript" onClick="save()"> <img class="Ico" src="images/save.gif" WIDTH="16" HEIGHT="16"> 
  </div>-->
  <div class="TBSep"></div>
</div>

<div class="yToolbar"> 
  <div class="TBHandle"> </div>
  <select id="FontName" class="TBGen" onChange="format('fontname',this[this.selectedIndex].value);this.selectedIndex=0">
    <option selected>����</option>
    <option value="����">����</option>
    <option value="����">����</option>
    <option value="����_GB2312">����</option>
    <option value="����_GB2312">����</option>
    <option value="����">����</option>
    <option value="��Բ">��Բ</option>
    <option value="Arial">Arial</option>
    <option value="Arial Black">Arial Black</option>
    <option value="Arial Narrow">Arial Narrow</option>
    <option value="Brush Script	MT">Brush Script MT</option>
    <option value="Century Gothic">Century Gothic</option>
    <option value="Comic Sans MS">Comic Sans MS</option>
    <option value="Courier">Courier</option>
    <option value="Courier New">Courier New</option>
    <option value="MS Sans Serif">MS Sans Serif</option>
    <option value="Script">Script</option>
    <option value="System">System</option>
    <option value="Times New Roman">Times New Roman</option>
    <option value="Verdana">Verdana</option>
    <option value="Wide	Latin">Wide Latin</option>
    <option value="Wingdings">Wingdings</option>
  </select>
  <select id="FontSize" class="TBGen" onChange="format('fontsize',this[this.selectedIndex].value);this.selectedIndex=0">
    <option selected>�ֺ�</option>
    <option value="7">һ��</option>
    <option value="6">����</option>
    <option value="5">����</option>
    <option value="4">�ĺ�</option>
    <option value="3">���</option>
    <option value="2">����</option>
    <option value="1">�ߺ�</option>
  </select>
  <div class="TBSep"></div>
  <div class="Btn" TITLE="�Ӵ�" LANGUAGE="javascript" onClick="format('bold')"> 
    <img class="Ico" src="images\bold.gif" WIDTH="16" HEIGHT="16"> </div>
  <div class="Btn" TITLE="б��" LANGUAGE="javascript" onClick="format('italic')"> 
    <img class="Ico" src="images\italic.gif" WIDTH="16" HEIGHT="16"> </div>
  <div class="Btn" TITLE="�»���" LANGUAGE="javascript" onClick="format('underline')"> 
    <img class="Ico" src="images\underline.gif" WIDTH="16" HEIGHT="16"> </div>
  <div class="TBSep"></div>
  <div class="Btn" TITLE="������ɫ" LANGUAGE="javascript" onClick="foreColor()"> <img class="Ico" src="images\fgcolor.gif" WIDTH="16" HEIGHT="16"> 
  </div>
  <div class="TBSep"></div>
  <div class="TBGen" title="�鿴HTMLԴ����"> 
    <input id="EditMode" onClick="setMode(this.checked)" type="checkbox">
    �鿴HTMLԴ����</div>
</div>

<iframe class="HtmlEdit" ID="HtmlEdit" MARGINHEIGHT="1" MARGINWIDTH="1" width="100%" height="320"> 
</iframe>
<script type="text/javascript">
SEP_PADDING = 5
HANDLE_PADDING = 7

var yToolbars =	new Array();
var YInitialized = false;
var bLoad=false
var pureText=true
var bodyTag="<head><style type=\"text/css\">body {font-size:	9pt}</style><meta http-equiv=Content-Type content=\"text/html; charset=gb2312\"></head><BODY bgcolor=\"#FFFFFF\" MONOSPACE>"
var bTextMode=false

public_description=new Editor

function document.onreadystatechange(){
  if (YInitialized) return;
  YInitialized = true;

  var i, s, curr;

  for (i=0; i<document.body.all.length;	i++)
  {
    curr=document.body.all[i];
    if (curr.className == "yToolbar")
    {
      InitTB(curr);
      yToolbars[yToolbars.length] = curr;
    }
  }

  DoLayout();
  window.onresize = DoLayout;

  HtmlEdit.document.open();
  HtmlEdit.document.write(bodyTag);
  HtmlEdit.document.close();
  HtmlEdit.document.designMode="On";
}

function InitBtn(btn)
{
  btn.onmouseover = BtnMouseOver;
  btn.onmouseout = BtnMouseOut;
  btn.onmousedown = BtnMouseDown;
  btn.onmouseup	= BtnMouseUp;
  btn.ondragstart = YCancelEvent;
  btn.onselectstart = YCancelEvent;
  btn.onselect = YCancelEvent;
  btn.YUSERONCLICK = btn.onclick;
  btn.onclick =	YCancelEvent;
  btn.YINITIALIZED = true;
  return true;
}

function InitTB(y)
{
  y.TBWidth = 0;

  if (!	PopulateTB(y)) return false;

  y.style.posWidth = y.TBWidth;

  return true;
}


function YCancelEvent()
{
  event.returnValue=false;
  event.cancelBubble=true;
  return false;
}

function PopulateTB(y)
{
  var i, elements, element;

  elements = y.children;
  for (i=0; i<elements.length; i++) {
    element = elements[i];
    if (element.tagName	== "SCRIPT" || element.tagName == "!") continue;

    switch (element.className) {
      case "Btn":
        if (element.YINITIALIZED == null)	{
          if (! InitBtn(element))
          return false;
        }
        element.style.posLeft = y.TBWidth;
        y.TBWidth	+= element.offsetWidth + 1;
        break;

      case "TBGen":
        element.style.posLeft = y.TBWidth;
        y.TBWidth	+= element.offsetWidth + 1;
        break;

      case "TBSep":
        element.style.posLeft = y.TBWidth	+ 2;
        y.TBWidth	+= SEP_PADDING;
        break;

      case "TBHandle":
        element.style.posLeft = 2;
        y.TBWidth	+= element.offsetWidth + HANDLE_PADDING;
        break;

      default:
        return false;
      }
  }

  y.TBWidth += 1;
  return true;
}

function DebugObject(obj)
{
  var msg = "";
  for (var i in	TB) {
    ans=prompt(i+"="+TB[i]+"\n");
    if (! ans) break;
  }
}

function LayoutTBs()
{
  NumTBs = yToolbars.length;

  if (NumTBs ==	0) return;

  var i;
  var ScrWid = (document.body.offsetWidth) - 6;
  var TotalLen = ScrWid;
  for (i = 0 ; i < NumTBs ; i++) {
    TB = yToolbars[i];
    if (TB.TBWidth > TotalLen) TotalLen	= TB.TBWidth;
  }

  var PrevTB;
  var LastStart	= 0;
  var RelTop = 0;
  var LastWid, CurrWid;
  var TB = yToolbars[0];
  TB.style.posTop = 0;
  TB.style.posLeft = 0;

  var Start = TB.TBWidth;
  for (i = 1 ; i < yToolbars.length ; i++) {
    PrevTB = TB;
    TB = yToolbars[i];
    CurrWid = TB.TBWidth;

    if ((Start + CurrWid) > ScrWid) {
      Start = 0;
      LastWid =	TotalLen - LastStart;
    }
    else {
       LastWid =	PrevTB.TBWidth;
       RelTop -=	TB.offsetHeight;
    }

    TB.style.posTop = RelTop;
    TB.style.posLeft = Start;
    PrevTB.style.width = LastWid;

    LastStart =	Start;
    Start += CurrWid;
  }

  TB.style.width = TotalLen - LastStart;

  i--;
  TB = yToolbars[i];
  var TBInd = TB.sourceIndex;
  var A	= TB.document.all;
  var item;
  for (i in A) {
    item = A.item(i);
    if (! item)	continue;
    if (! item.style) continue;
    if (item.sourceIndex <= TBInd) continue;
    if (item.style.position == "absolute") continue;
    item.style.posTop =	RelTop;
  }
}

function DoLayout()
{
  LayoutTBs();
}

function BtnMouseOver()
{
  if (event.srcElement.tagName != "IMG") return	false;
  var image = event.srcElement;
  var element =	image.parentElement;

  if (image.className == "Ico")	element.className = "BtnMouseOverUp";
  else if (image.className == "IcoDown") element.className = "BtnMouseOverDown";

  event.cancelBubble = true;
}

function BtnMouseOut()
{
  if (event.srcElement.tagName != "IMG") {
    event.cancelBubble = true;
    return false;
  }

  var image = event.srcElement;
  var element =	image.parentElement;
  yRaisedElement = null;

  element.className = "Btn";
  image.className = "Ico";

  event.cancelBubble = true;
}

function BtnMouseDown()
{
  if (event.srcElement.tagName != "IMG") {
    event.cancelBubble = true;
    event.returnValue=false;
    return false;
  }

  var image = event.srcElement;
  var element =	image.parentElement;

  element.className = "BtnMouseOverDown";
  image.className = "IcoDown";

  event.cancelBubble = true;
  event.returnValue=false;
  return false;
}

function BtnMouseUp()
{
  if (event.srcElement.tagName != "IMG") {
    event.cancelBubble = true;
    return false;
  }

  var image = event.srcElement;
  var element =	image.parentElement;

  if(navigator.appVersion.match(/8./i)=='8.') 
    { 
      if (element.YUSERONCLICK) eval(element.YUSERONCLICK + "onclick(event)");  
  } 
else 

  { 
    if (element.YUSERONCLICK) eval(element.YUSERONCLICK + "anonymous()"); 
} 

  element.className = "BtnMouseOverUp";
  image.className = "Ico";

  event.cancelBubble = true;
  return false;
}

function getEl(sTag,start)
{
  while	((start!=null) && (start.tagName!=sTag)) start = start.parentElement;
  return start;
}

function cleanHtml()
{
  var fonts = HtmlEdit.document.body.all.tags("FONT");
  var curr;
  for (var i = fonts.length - 1; i >= 0; i--) {
    curr = fonts[i];
    if (curr.style.backgroundColor == "#ffffff") curr.outerHTML	= curr.innerHTML;
  }
}

function getPureHtml()
{
  var str = "";
  var paras = HtmlEdit.document.body.all.tags("P");
  if (paras.length > 0)	{
    for	(var i=paras.length-1; i >= 0; i--) str	= paras[i].innerHTML + "\n" + str;
  }
  else {
    str	= HtmlEdit.document.body.innerHTML;
  }
  return str;
}


function Editor()
{
  this.put_HtmlMode=setMode;
  this.put_value=putText;
  this.get_value=getText;
}

function getText()
{
  if (bTextMode)
    return HtmlEdit.document.body.innerText;
  else
  {
    cleanHtml();
    cleanHtml();
    return HtmlEdit.document.body.innerHTML;
  }
}

function putText(v)
{
  if (bTextMode)
    HtmlEdit.document.body.innerText = v;
  else
    HtmlEdit.document.body.innerHTML = v;
}

function UserDialog(what)
{
  if (!validateMode()) return;

  HtmlEdit.document.execCommand(what, true);

  pureText = false;
  HtmlEdit.focus();
}

function validateMode()
{
  if (!	bTextMode) return true;
  alert("��ȡ�����鿴HTMLԴ���롱ѡ�Ȼ����ʹ��ϵͳ�༭����!");
  HtmlEdit.focus();
  return false;
}

function format(what,opt)
{
  if (!validateMode()) return;
  if (opt=="removeFormat")
  {
    what=opt;
    opt=null;
  }

  if (opt==null) HtmlEdit.document.execCommand(what);
  else HtmlEdit.document.execCommand(what,"",opt);

  pureText = false;
  HtmlEdit.focus();
}

function setMode(newMode)
{
  var cont;
  bTextMode = newMode;
  if (bTextMode) {
    cleanHtml();
    cleanHtml();

    cont=HtmlEdit.document.body.innerHTML;
    HtmlEdit.document.body.innerText=cont;
  }
  else {
    cont=HtmlEdit.document.body.innerText;
    HtmlEdit.document.body.innerHTML=cont;
  }
  HtmlEdit.focus();
}

function foreColor()
{
  if (!	validateMode())	return;
  var arr = showModalDialog("selcolor.asp", "", "dialogWidth:18.5em; dialogHeight:17.5em; status:0");
  if (arr != null) format('forecolor', arr);
  else HtmlEdit.focus();
}

function InsertTable()
{
  if (!	validateMode())	return;
  HtmlEdit.focus();
  var range = HtmlEdit.document.selection.createRange();
  var arr = showModalDialog("table.asp", "", "dialogWidth:300pt;dialogHeight:236pt;help:0;status:0");

  if (arr != null){
	range.pasteHTML(arr);
  }
  HtmlEdit.focus();
}


function InsertImg()
{
  if (!	validateMode())	return;
  HtmlEdit.focus();
  var range = HtmlEdit.document.selection.createRange();
  var arr = showModalDialog("image.asp", "", "dialogWidth:430px; dialogHeight:230px; status:0");
  if (arr != null)
  {
	range.pasteHTML(arr);
	parent.myform.IncludePic.checked=true;
  }
  HtmlEdit.focus();
}



function specialtype(Mark){
  if (!Error()) return;
  var sel,RangeType
  sel = HtmlEdit.document.selection.createRange();
  RangeType = HtmlEdit.document.selection.type;
  if (RangeType == "Text"){
    sel.pasteHTML("<" + Mark + ">" + sel.text + "</" + Mark + ">");
    sel.select();
  }
  HtmlEdit.focus();
}

function help()
{
  var arr = showModalDialog("help.asp", "", "dialogWidth:580px; dialogHeight:460px; status:0");
}

function save()
{
  if (bTextMode){
//�༭��Ƕ��������ҳʱʹ��������һ�䣨�뽫form1�ĳ���Ӧ��������
  parent.myform.Profile.value=HtmlEdit.document.body.innerText;
//�����򿪱༭��ʱʹ��������һ�䣨�뽫form1�ĳ���Ӧ��������  
//  self.opener.form1.content.value+=HtmlEdit.document.body.innerText;
  }
  else{
//�༭��Ƕ��������ҳʱʹ��������һ�䣨�뽫form1�ĳ���Ӧ��������
  parent.myform.Profile.value=HtmlEdit.document.body.innerHTML;
//�����򿪱༭��ʱʹ��������һ�䣨�뽫form1�ĳ���Ӧ��������  
//  self.opener.form1.content.value+=HtmlEdit.document.body.innerHTML;
  }
  HtmlEdit.focus();
  return false;
}
</script>

</body>
</html>