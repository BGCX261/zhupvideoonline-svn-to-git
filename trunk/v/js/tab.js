//tabҳ,���ò�ʵ��.���Ըı�TABҳ�������ɫ(tablayeraΪ���ݲ�,tabaΪtab��id,��˿��Ե���任��ʽ)
function showlaya(i){

var tablayera = document.getElementById("tablayera" + i);
var taba = document.getElementById("taba" + i);
var tabpic = document.getElementById("tabpic" + i);

for(var j=1; j<3; j++){
	var tablayera1 = document.getElementById("tablayera" + j);
	var taba1 = document.getElementById("taba" + j);
	var tabpic1 = document.getElementById("tabpic" + j);
	if(tablayera1 == undefined){
		break;
	}else{
		tablayera1.style.display="none";
		taba1.className="tab_titledown";
		tabpic1.style.display="none";
	}
} 

tablayera.style.display="block";
taba.className="tab_titleup";
tabpic.style.display="block";

}