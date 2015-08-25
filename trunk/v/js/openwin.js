//留言本相关脚本
var oldContent="";

function turnit(newContent) {

  if(document.all[newContent].style.display=="none") {
    if(oldContent!="") {
      document.all[oldContent].style.display="none";
    }
    oldContent=newContent;
    document.all[newContent].style.display="";
  } else {
    if(oldContent!="")
    oldContent=newContent;
    document.all[newContent].style.display="none"; 
  }
}

function turnit2() {
  if(oldContent!="") {
    document.all[oldContent].style.display="none";
    oldContent="";
  }
}
