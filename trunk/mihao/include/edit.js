function Onclicklink() 
{ 
var url = spreadurl.innerHTML
window.clipboardData.setData("Text",url);
alert("�Ѹ�������");
} 

//��Ա����
function openjifen(id)
{
	var x,y;
	x=(screen.width-640)/3;
	y=(screen.height-480)/3;
	window.open("jifen.asp?id="+id,"memberwindow","toolbar=no,resizable=no,directories=no,status=yes,scrollbars=auto,menubar=no,width=330,height=200,left="+x+",top="+y);
}

//ע�ᴰ��
function openreg()
{
	var x,y;
	x=(screen.width-640)/3;
	y=(screen.height-480)/3;
	window.open("register.asp","regwindow","toolbar=no,resizable=no,directories=no,status=yes,scrollbars=auto,menubar=no,width=380,height=280,left="+x+",top="+y);
}

//��Ա��Ϣ
function memberinfo(id)
{
	var x,y;
	x=(screen.width-640)/3;
	y=(screen.height-480)/3;
	window.open("memberinfo.asp?id="+id,"memberwindow","toolbar=no,resizable=no,directories=no,status=yes,scrollbars=auto,menubar=no,width=380,height=350,left="+x+",top="+y);
}

function memberinfo2(id)
{
	var x,y;
	x=(screen.width-640)/3;
	y=(screen.height-480)/3;
	window.open("../memberinfo.asp?id="+id,"memberwindow","toolbar=no,resizable=no,directories=no,status=yes,scrollbars=auto,menubar=no,width=400,height=380,left="+x+",top="+y);
}


//��¼����
function openlogin(id)
{
	var x,y;
	x=(screen.width-640)/3;
	y=(screen.height-480)/3;
	window.open("include/logins.asp","loginwindow","toolbar=no,resizable=no,directories=no,status=yes,scrollbars=auto,menubar=no,width=320,height=260,left="+x+",top="+y);
}

//�����ȴ�
function openhotkw(id)
{
	var x,y;
	x=(screen.width-640)/3;
	y=(screen.height-480)/3;
	window.open("admin_config.asp?id="+id,"memberwindow","toolbar=no,resizable=no,directories=no,status=yes,scrollbars=auto,menubar=no,width=850,height=138,left="+x+",top="+y);
}

//���ٷ�����Ϣ
function opennotice(xx)
{
	var x,y;
	x=(screen.width-640)/3;
	y=(screen.height-480)/3;
	window.open("admin_show.asp?target=notice&xx="+xx,"memberwindow","toolbar=no,resizable=no,directories=no,status=yes,scrollbars=auto,menubar=no,width=850,height=138,left="+x+",top="+y);
}

//�������
function DoTitle(addTitle) { 
var revisedTitle; 
var currentTitle = document.form1.title.value; 
addTitle="["+addTitle+"]"
revisedTitle = addTitle+String.fromCharCode(32)+currentTitle; 
document.form1.title.value=revisedTitle; 
document.form1.title.focus();
return; 
}

//����ظ�����
function setReplyTitle(title,idnum) {
var newtitle;
newtitle="�ظ�"+idnum+"¥:"+title
if (idnum==0)
{
newtitle="�ظ�¥��:"+title;
}
document.form1.title.value=newtitle;
checkFocus();
return;
}

//����ͼƬ�Զ�������ȴ�С����ֹ����
flag=false
function DrawImage(ImgD){ 
	var image=new Image(); 
	image.src=ImgD.src; 
	if(image.width>0 && image.height>0){ 
		flag=true; 
		if(image.width>=580){ 
			ImgD.width=580; 
			ImgD.height=(image.height*580)/image.width; 
		}else{ 
			ImgD.width=image.width; 
			ImgD.height=image.height; 
		}  
	} 
} 

//����ͼƬ����
function errpic(thepic){
thepic.src="images/showerr.gif" 
}

//���ݻ���
function CopyText(obj) {
	ie = (document.all)? true:false
	if (ie){
		var rng = document.body.createTextRange();
		rng.moveToElementText(obj);
		rng.scrollIntoView();
		rng.select();
		rng.execCommand("Copy");
		rng.collapse(false);
	}
}

//ҳ��ؼ�������
function searchkey(objname,num){
	var myfontsize,a,i;
	var tempText,Text;
	for (i=1;i<=num ;i++ )
	{
	 a=objname+i;
	 //document.write(a);
	Text=document.getElementById(a).innerText;
	tempText=Text;
	h="<font class=keyword>";
	f="</font>";
	keyset=new Array();
	key=document.all.keys.value;
	if (key==""){
		alert("�����������ؼ���!");
			return;
		}
		else{
		keyset[0]=tempText.indexOf(key,0);
			if (keyset[0]<0){
					return;
			}else
				temp=tempText.substring(0,keyset[0]);
				temp=temp+h+key+f;
				temp2=tempText.substring(keyset[0]+key.length,tempText.length);
				for (i=1;i<tempText.length;i++)	{
					keyset[i]=tempText.indexOf(key,keyset[i-1]+key.length);
					if(keyset[i]<0){
					temp=temp+tempText.substring(keyset[i-1]+key.length,tempText.length);
					break;
					}else{
					temp=temp+tempText.substring(keyset[i-1]+key.length,keyset[i])+h+key+f;
					}
				}
					a.innerHTML = temp;
			}
	}
	}


/*����*/
function quoteit(objname) {
		if (document.selection && document.selection.type == "Text") {
		var range = document.selection.createRange();
		if(wysiwyg)	
			{editdoc.body.innerHTML+="[quote][link="+objname+"][/link][color=gray]" + document.getElementById(objname).innerText + "[/color]<br/>===============================================<br/>" + range.text + "[/quote]\n";
			}
		else 
			{$('message').value +="[quote][link="+objname+"][/link][color=gray]" + document.getElementById(objname).innerText + "[/color]\n===============================================\n" + range.text + "[/quote]\n";
				}
		alert("���������ѳɹ����Ƶ���������");
		checkFocus();	
		} else {  
        alert("����������϶�ѡ��Ҫ���õ�����!");
        }
}


/*��дCookie*/

function readCookie(name)
{
  var cookieValue = "";
  var search = name + "=";
  if(document.cookie.length > 0)
  { 
    offset = document.cookie.indexOf(search);
    if (offset != -1)
    { 
      offset += search.length;
      end = document.cookie.indexOf(";", offset);
      if (end == -1) end = document.cookie.length;
      cookieValue = unescape(document.cookie.substring(offset, end))
    }
  }
  return cookieValue;
}

// Example:
// writeCookie("myCookie", "my name", 24);
// Stores the string "my name" in the cookie "myCookie" which expires after 24 hours.
function writeCookie(name, value, hours)
{
  var expire = "";
  if(hours != null)
  {
    expire = new Date((new Date()).getTime() + hours * 3600000);
    expire = "; expires=" + expire.toGMTString();
  }
  document.cookie = name + "=" + escape(value) + expire;
  
}	

function askdelete()
{
	if(confirm("��ȷ��Ҫɾ��������⣨����һ���ظ�����"))
		return(true);
	else
		return(false);
}

