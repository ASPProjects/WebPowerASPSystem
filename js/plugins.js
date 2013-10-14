// JavaScript Document
//IVY page plugins

//menu plugins
var nw=0;
function menuEffect(status,obj){
	var strdiv=obj.getElementsByTagName("div").item(0);
	var slength=strdiv.getElementsByTagName("a").length;

	var strheight=slength*24+11;
	if (status=="show"){obj.className="on";
	//strdiv.style.height=strheight+"px";
	strdiv.style.display="block";
	//window.setInterval(menusub,1000); 
	//menusub();
	}
	if (status=="hidden"){obj.className="";
	strdiv.style.display="none";}
	//alert(obj.className);
	
function menusub(){
	if (nw<slength){	
	alert(strdiv.style.height);
	strdiv.style.height=nw+"px";
	nw=nw+2;
	}else{	
	strdiv.style.height=strheight+"px";
	}
	}
}
