//最后修订版本号:$Revision: 28 $
//最后修改人:$Author: admin $
//最后修改时间:$Date: 2009-07-20 18:03:52 +0800 (星期一, 2009-07-20) $
/*********************************************************************************************/
/*功能:JQ插件图片预载入等比居中缩放*/
/*********************************************************************************************/
jQuery.fn.LoadImage=function(width,height,txtheight,loadpic){
  if(loadpic==null)loadpic="/Upfiles/system/loading.gif";
  return this.each(function(){
    var t=$(this);
    //取父框div.img的内外框值
    var	parentPadTop=Math.round(t.parent().css("padding-top").replace(/\D/g,''));
    var	parentPadBom=Math.round(t.parent().css("padding-bottom").replace(/\D/g,''));
    var	parentMarBom=Math.round(t.parent().css("margin-bottom").replace(/\D/g,''));
    var	parentMarTop=Math.round(t.parent().css("margin-top").replace(/\D/g,''));
    var	parentPadLft=Math.round(t.parent().css("padding-left").replace(/\D/g,''));
    var	parentPadRgt=Math.round(t.parent().css("padding-right").replace(/\D/g,''));
    var	parentMarLft=Math.round(t.parent().css("margin-left").replace(/\D/g,''));
    var	parentMarRgt=Math.round(t.parent().css("margin-right").replace(/\D/g,''));
    var	parentBdRgt=Math.round(t.parent().css("border-right-width").replace(/\D/g,''));
    var	parentBdLft=Math.round(t.parent().css("border-left-width").replace(/\D/g,''));
    var	parentBdTop=Math.round(t.parent().css("border-top-width").replace(/\D/g,''));
    var	parentBdBom=Math.round(t.parent().css("border-bottom-width").replace(/\D/g,''));
    var	boxWidth=parentBdLft+parentPadLft+parentMarLft+parentBdRgt+parentPadRgt+parentMarRgt;
    var	boxHeight=parentBdTop+parentPadTop+parentMarTop+parentBdBom+parentPadBom+parentMarBom;
    //设定父级外框a元素高度
    var aheight=height+txtheight;		
    t.parent().parent().css({height:aheight,width:width});
    var src=$(this).attr("src");
    var img=new Image();
    //alert("Loading...")
    img.src=src;
    //自动缩放图片
    var autoScaling=function(){
      
      
        if(img.width>0 && img.height>0){ 
              if(img.width/img.height>=width/height){ 
                  if(img.width>width){ 
                      t.width(width-boxWidth); 
                      t.height((img.height*width)/img.width-boxHeight); 
                  }else{ 
                      t.width(img.width-boxWidth); 
                      t.height(img.height-boxHeight); 
                  } 
              } 
              else{ 
                  if(img.height>height){ 
                      t.height(height-boxWidth); 
                      t.width((img.width*height)/img.height-boxHeight); 
                  }else{ 
                      t.width(img.width-boxWidth); 
                      t.height(img.height-boxHeight); 
                  } 
              } 
          } 
        
        if (t.height()<height){
          //取父padding和margin值

          var temppad=Math.abs(Math.round((height-t.height()-boxWidth)/2));
          t.css({"margin-top":temppad});
          t.css({"margin-bottom":temppad});
        }
        if (t.width()<width){

          var temppad2=Math.abs(Math.round((width-t.width()-boxHeight)/2));
          t.css({"margin-left":temppad2});
          t.css({"margin-right":temppad2});
        }

      }	
  
    //处理ff下会自动读取缓存图片
    if(img.complete){
        //alert("getToCache!");
      autoScaling();
        return;
    }
    $(this).attr("src","");
    var loading=$("<img alt=\"加载中...\" title=\"图片加载中...\" src=\""+loadpic+"\" />");
    
    t.hide();
    t.after(loading);
    $(img).load(function(){
      autoScaling();
      loading.remove();
      t.attr("src",this.src);
      t.css()
      t.show();
      //alert("finally!")
    });
    
  });
}
/*********************************************************************************************/
/*功能:JQ插件页面漂浮在线QQ之类*/
/*********************************************************************************************/
jQuery.fn.jFloat = function(o) {
    
        o = $.extend({
            top:60,  //广告距页面顶部距离
            left:0,//广告左侧距离
            right:0,//广告右侧距离
            width:100,  //广告容器的宽度
            height:360, //广告容器的高度
            minScreenW:800,//出现广告的最小屏幕宽度，当屏幕分辨率小于此，将不出现对联广告
            position:"left", //对联广告的位置left-在左侧出现,right-在右侧出现
            allowClose:true //是否允许关闭 
        }, o || {});
    var h=o.height;
      var showAd=true;
      var fDiv=$(this);
      if(o.minScreenW>=$(window).width()){
          fDiv.hide();
          showAd=false;
       }
       else{
       fDiv.css("display","block")
           var closeHtml='<div align="right" style="padding:2px;z-index:2000;font-size:12px;cursor:pointer;border-bottom:1px solid #f1f1f1; height:20px;" class="closeFloat"><span style="border:1px solid #000;height:12px;display:block;width:12px;">×</span></div>';
           switch(o.position){
               case "left":
                    if(o.allowClose){
                       fDiv.prepend(closeHtml);
             $(".closeFloat",fDiv).click(function(){$(this).hide();fDiv.hide();showAd=false;})
             h+=20;
          }
                    fDiv.css({position:"absolute",left:o.left+"px",top:o.top+"px",width:o.width+"px",height:h+"px",overflow:"hidden"});
                    break;
               case "right":
                    if(o.allowClose){
                       fDiv.prepend(closeHtml)
             $(".closeFloat",fDiv).click(function(){$(this).hide();fDiv.hide();showAd=false;})
             h+=20;
          }
                    fDiv.css({position:"absolute",left:"auto",right:o.right+"px",top:o.top+"px",width:o.width+"px",height:h+"px",overflow:"hidden"});
                    break;
            };
        };
        function ylFloat(){
            if(!showAd){return}
            var windowTop=$(window).scrollTop();
            if(fDiv.css("display")!="none")
                fDiv.css("top",o.top+windowTop+"px");
        };

      $(window).scroll(ylFloat) ;
      $(document).ready(ylFloat);     
       
    }
/*********************************************************************************************/
/*功能:JQ插件横向滚动使用*/
/*********************************************************************************************/
jQuery.fn.jMarquee = function(o) {
    o = $.extend({
    speed:30,
    step:1,//滚动步长
    direction:"up",//滚动方向
    visible:1//可见元素数量
    }, o || {});
    //获取滚动内容内各元素相关信息
    var i=0;
    var div=$(this);
    var ul=$("ul",div);
    var tli=$("li",ul);
    var liSize=tli.size();
    if(o.direction=="left")
        tli.css("float","left");
    var liWidth=tli.innerWidth();
    var liHeight=tli.height();
    var ulHeight=liHeight*liSize;
    var ulWidth=liWidth*liSize;
  
    //如果对象元素个数大于指定的显示元素则进行滚动，否则不滚动。
    if(liSize>o.visible){
        ul.append(tli.slice(0,o.visible).clone())  //复制前o.visible个li，并添加到ul的最后
        li=$("li",ul);
        liSize=li.size();
        
          //给滚动内容添加相关CSS样式
        div.css({"position":"relative",overflow:"hidden"});
        ul.css({"position":"relative",margin:"0",padding:"0","list-style":"none"});
        li.css({margin:"0",padding:"0","position":"relative"});
        
        switch(o.direction){
          case "left":
              div.css("width",(liWidth*o.visible)+"px");
              ul.css("width",(liWidth*liSize)+"px");
              li.css("float","left");
              break;
          case "up":
              div.css({"height":(liHeight*o.visible)+"px"});
              ul.css("height",(liHeight*liSize)+"px");
              break;
        }
        
       
        var MyMar=setInterval(ylMarquee,o.speed);
        ul.hover(
            function(){clearInterval(MyMar);},
            function(){MyMar=setInterval(ylMarquee,o.speed);}
        );
    };
    function ylMarquee(){
         
        if(o.direction=="left"){
            if(div.scrollLeft()>=ulWidth){
                div.scrollLeft(0);
            }
            else
            {
                var leftNum=div.scrollLeft();
                leftNum+=parseInt(o.step);
                div.scrollLeft(leftNum)
            }
        }
        
        if(o.direction=="up"){
            if(div.scrollTop()>=ulHeight){
               div.scrollTop(0);
                
            }
            else{
               var topNum=div.scrollTop();
               topNum+=parseInt(o.step);
               div.scrollTop(topNum);
            }
        }
        
    };
   
}
/*********************************************************************************************/
/*功能:JQ插件页面等高使用*/
/*********************************************************************************************/
$.fn.equalHeights = function(px) {
  $(this).each(function(){
    var currentTallest = 0;
    $(this).children().each(function(i){
      if ($(this).height() > currentTallest) { currentTallest = $(this).height(); }
    });
    if (!px || !Number.prototype.pxToEm) currentTallest = currentTallest.pxToEm(); //use ems unless px is specified
    // for ie6, set height since min-height isn't supported
    if ($.browser.msie && $.browser.version == 6.0) { $(this).children().css({'height': currentTallest}); }
    $(this).children().css({'min-height': currentTallest}); 
  });
  return this;
};

// just in case you need it...
$.fn.equalWidths = function(px) {
  $(this).each(function(){
    var currentWidest = 0;
    $(this).children().each(function(i){
        if($(this).width() > currentWidest) { currentWidest = $(this).width(); }
    });
    if(!px || !Number.prototype.pxToEm) currentWidest = currentWidest.pxToEm(); //use ems unless px is specified
    // for ie6, set width since min-width isn't supported
    if ($.browser.msie && $.browser.version == 6.0) { $(this).children().css({'width': currentWidest}); }
    $(this).children().css({'min-width': currentWidest}); 
  });
  return this;
};
Number.prototype.pxToEm = String.prototype.pxToEm = function(settings){
  //set defaults
  settings = jQuery.extend({
    scope: 'body',
    reverse: false
  }, settings);
  
  var pxVal = (this == '') ? 0 : parseFloat(this);
  var scopeVal;
  var getWindowWidth = function(){
    var de = document.documentElement;
    return self.innerWidth || (de && de.clientWidth) || document.body.clientWidth;
  };	
        
  if (settings.scope == 'body' && $.browser.msie && (parseFloat($('body').css('font-size')) / getWindowWidth()).toFixed(1) > 0.0) {
    var calcFontSize = function(){		
      return (parseFloat($('body').css('font-size'))/getWindowWidth()).toFixed(3) * 16;
    };
    scopeVal = calcFontSize();
  }
  else { scopeVal = parseFloat(jQuery(settings.scope).css("font-size")); };
      
  var result = (settings.reverse == true) ? (pxVal * scopeVal).toFixed(2) + 'px' : (pxVal / scopeVal).toFixed(2) + 'em';
  return result;
};

/*********************************************************************************************/
/*功能:JQ插件,幻灯图片*/
/*********************************************************************************************/
jQuery.fn.focusImg=function(){
 var obj=$(this);
 objName=obj.attr("id");
 obj.children("ul").attr("id",objName+"-fragment");
 var num=(obj.children("ul").children("li").length)-1;
 var temp_b=new String();
 for(var i=0;i<=num;i++){
  temp_b+="<a id=\""+objName+"-button-"+i+"\" href=\"#"+objName+"-fragment-"+i+"\">"+(i+1)+"</a>";
  obj.children("ul").children("li").eq(i).attr("id",objName+"-fragment-"+i)
 }
 obj.append("<span id='"+objName+"-button'>"+temp_b+"</span>");
 obj.children("span").children("a").eq(num).addClass("a2");

 obj.children("span").children("a").click(function(){
   obj.children("span").children("a").removeClass("a2");
   $(this).addClass("a2");
   var id=($(this).attr("id")).replace("button","fragment");
   $("#"+id).appendTo(obj.children("ul"));
   return false;
 });

 topBoxRun=function(){
  obj.children("ul").children("li").eq(num).fadeOut("slow",function(){
   obj.children("span").children("a").removeClass("a2");
   var id=($(this).prev().attr("id")).replace("fragment","button");
   $("#"+id).addClass("a2");
   $(this).prependTo(obj.children("ul"));
   $(this).show();
  }); 
 }

 var t=setInterval(topBoxRun,3000);
}

/*********************************************************************************************/
/*功能:JQ插件,横向竖直两用多级ul li形菜单*/
/*********************************************************************************************/
var ddsmoothmenu={

//Specify full URL to down and right arrow images (23 is padding-right added to top level LIs with drop downs):
arrowimages: {down:['downarrowclass', 'down.gif', 23], right:['rightarrowclass', 'right.gif']},

transition: {overtime:300, outtime:300}, //duration of slide in/ out animation, in milliseconds
shadow: {enabled:true, offsetx:5, offsety:5},

///////Stop configuring beyond here///////////////////////////

detectwebkit: navigator.userAgent.toLowerCase().indexOf("applewebkit")!=-1, //detect WebKit browsers (Safari, Chrome etc)
detectie6: document.all && !window.XMLHttpRequest,

getajaxmenu:function($, setting){ //function to fetch external page containing the panel DIVs
  var $menucontainer=$('#'+setting.contentsource[0]) //reference empty div on page that will hold menu
  $menucontainer.html("Loading Menu...")
  $.ajax({
    url: setting.contentsource[1], //path to external menu file
    async: true,
    error:function(ajaxrequest){
      $menucontainer.html('Error fetching content. Server Response: '+ajaxrequest.responseText)
    },
    success:function(content){
      $menucontainer.html(content)
      ddsmoothmenu.buildmenu($, setting)
    }
  })
},


buildmenu:function($, setting){
  var smoothmenu=ddsmoothmenu
  var $mainmenu=$("#"+setting.mainmenuid+">ul") //reference main menu UL
  $mainmenu.parent().get(0).className=setting.classname || "ddsmoothmenu"
  var $headers=$mainmenu.find("ul").parent()
  $headers.hover(
    function(e){
      $(this).children('a:eq(0)').addClass('selected')
    },
    function(e){
      $(this).children('a:eq(0)').removeClass('selected')
    }
  )
  $headers.each(function(i){ //loop through each LI header
    var $curobj=$(this).css({zIndex: 100-i}) //reference current LI header
    var $subul=$(this).find('ul:eq(0)').css({display:'block'})
    this._dimensions={w:this.offsetWidth, h:this.offsetHeight, subulw:$subul.outerWidth(), subulh:$subul.outerHeight()}
    this.istopheader=$curobj.parents("ul").length==1? true : false //is top level header?
    $subul.css({top:this.istopheader && setting.orientation!='v'? this._dimensions.h+"px" : 0})
    $curobj.children("a:eq(0)").css(this.istopheader? {paddingRight: smoothmenu.arrowimages.down[2]} : {}).append( //add arrow images
      '<img src="'+ (this.istopheader && setting.orientation!='v'? smoothmenu.arrowimages.down[1] : smoothmenu.arrowimages.right[1])
      +'" class="' + (this.istopheader && setting.orientation!='v'? smoothmenu.arrowimages.down[0] : smoothmenu.arrowimages.right[0])
      + '" style="border:0;" />'
    )
    if (smoothmenu.shadow.enabled){
      this._shadowoffset={x:(this.istopheader?$subul.offset().left+smoothmenu.shadow.offsetx : this._dimensions.w), y:(this.istopheader? $subul.offset().top+smoothmenu.shadow.offsety : $curobj.position().top)} //store this shadow's offsets
      if (this.istopheader)
        $parentshadow=$(document.body)
      else{
        var $parentLi=$curobj.parents("li:eq(0)")
        $parentshadow=$parentLi.get(0).$shadow
      }
      this.$shadow=$('<div class="ddshadow'+(this.istopheader? ' toplevelshadow' : '')+'"></div>').prependTo($parentshadow).css({left:this._shadowoffset.x+'px', top:this._shadowoffset.y+'px'})  //insert shadow DIV and set it to parent node for the next shadow div
    }
    $curobj.hover(
      function(e){
        var $targetul=$(this).children("ul:eq(0)")
        this._offsets={left:$(this).offset().left, top:$(this).offset().top}
        var menuleft=this.istopheader && setting.orientation!='v'? 0 : this._dimensions.w
        menuleft=(this._offsets.left+menuleft+this._dimensions.subulw>$(window).width())? (this.istopheader && setting.orientation!='v'? -this._dimensions.subulw+this._dimensions.w : -this._dimensions.w) : menuleft //calculate this sub menu's offsets from its parent
        if ($targetul.queue().length<=1){ //if 1 or less queued animations
          $targetul.css({left:menuleft+"px", width:this._dimensions.subulw+'px'}).animate({height:'show',opacity:'show'}, ddsmoothmenu.transition.overtime)
          if (smoothmenu.shadow.enabled){
            var shadowleft=this.istopheader? $targetul.offset().left+ddsmoothmenu.shadow.offsetx : menuleft
            var shadowtop=this.istopheader?$targetul.offset().top+smoothmenu.shadow.offsety : this._shadowoffset.y
            if (!this.istopheader && ddsmoothmenu.detectwebkit){ //in WebKit browsers, restore shadow's opacity to full
              this.$shadow.css({opacity:1})
            }
            this.$shadow.css({overflow:'', width:this._dimensions.subulw+'px', left:shadowleft+'px', top:shadowtop+'px'}).animate({height:this._dimensions.subulh+'px'}, ddsmoothmenu.transition.overtime)
          }
        }
      },
      function(e){
        var $targetul=$(this).children("ul:eq(0)")
        $targetul.animate({height:'hide', opacity:'hide'}, ddsmoothmenu.transition.outtime)
        if (smoothmenu.shadow.enabled){
          if (ddsmoothmenu.detectwebkit){ //in WebKit browsers, set first child shadow's opacity to 0, as "overflow:hidden" doesn't work in them
            this.$shadow.children('div:eq(0)').css({opacity:0})
          }
          this.$shadow.css({overflow:'hidden'}).animate({height:0}, ddsmoothmenu.transition.outtime)
        }
      }
    ) //end hover
  }) //end $headers.each()
  $mainmenu.find("ul").css({display:'none', visibility:'visible'})
},

init:function(setting){
  if (typeof setting.customtheme=="object" && setting.customtheme.length==2){ //override default menu colors (default/hover) with custom set?
    var mainmenuid='#'+setting.mainmenuid
    var mainselector=(setting.orientation=="v")? mainmenuid : mainmenuid+', '+mainmenuid
    document.write('<style type="text/css">\n'
      +mainselector+' ul li a {background:'+setting.customtheme[0]+';}\n'
      +mainmenuid+' ul li a:hover {background:'+setting.customtheme[1]+';}\n'
    +'</style>')
  }
  this.shadow.enabled=(document.all && !window.XMLHttpRequest)? false : true //in IE6, always disable shadow
  jQuery(document).ready(function($){ //ajax menu?
    if (typeof setting.contentsource=="object"){ //if external ajax menu
      ddsmoothmenu.getajaxmenu($, setting)
    }
    else{ //else if markup menu
      ddsmoothmenu.buildmenu($, setting)
    }
  })
}

} //end ddsmoothmenu variable

//Initialize Menu instance(s):
/*********************************************************************************************/
/*功能:JQ插件,轻便ul li下拉菜单*/
/*********************************************************************************************/
jQuery.jFastMenu = function(id){

  $(id + ' ul li').hover(function(){
    $(this).find('ul:first').animate({height:'show'}, 'fast');
  },
  function(){
    $(this).find('ul:first').animate({height:'hide', opacity:'hide'}, 'slow');
  });

}

/*********************************************************************************************/
/*功能:折叠菜单使用*/
/*********************************************************************************************/
function ulshow(vid)
{
if(document.getElementById("ul"+vid).style.display=="none"){
   document.getElementById("ul"+vid).style.display="";
   document.getElementById("f"+vid).src=document.getElementById("f"+vid).src.replace("+","-")
}
else{
   document.getElementById("ul"+vid).style.display="none";
   document.getElementById("f"+vid).src=document.getElementById("f"+vid).src.replace("-","+")
}
return;
}

/*********************************************************************************************/
/*功能:搜索页使用*/
/*********************************************************************************************/
function getweblist2(url,domid,pagenum){
var CurrentFunPath;
var sUSERNAME;
var sGUID;
if(typeof(VideoPathRoot)=="undefined"){CurrentFunPath=""}else{CurrentFunPath=VideoPathRoot}
if(typeof(VideoUSERNAME)=="undefined"){sUSERNAME=""}else{sUSERNAME=VideoUSERNAME}
if(typeof(VideoGUID)=="undefined"){sGUID=""}else{sGUID=VideoGUID}
var objectn="#"+domid;
var targeturl="http://www.51g3.com/news/remote3.0.asp?callback=?&page="+pagenum+"&pagenum="+url+"&ntimest="+Math.random()+"&guid="+sGUID+"&u="+sUSERNAME;
 $.getJSON(targeturl,function(d){
    $(objectn).html(d.jsonObj[0]);
  });
}


function autoplay(url,pagenum){
var CurrentFunPath;
var sUSERNAME;
var sGUID;
if(typeof(VideoPathRoot)=="undefined"){CurrentFunPath=""}else{CurrentFunPath=VideoPathRoot}
if(typeof(VideoUSERNAME)=="undefined"){sUSERNAME=""}else{sUSERNAME=VideoUSERNAME}
if(typeof(VideoGUID)=="undefined"){sGUID=""}else{sGUID=VideoGUID}
var VideoBaseURL;
if (CurrentFunPath=="") 
{VideoBaseURL=CurrentFunPath}
else
{
if(CurrentFunPath.indexOf('://')>=0)
{VideoBaseURL = CurrentFunPath;
}
else
{VideoBaseURL = 'http://'+window.location.host+'/';}
}
var targeturl=""+VideoBaseURL+"inc/ajaxtv.asp?callback=?&page="+pagenum+"&pagenum="+url+"&ntimest="+Math.random()+"&guid="+sGUID+"&u="+sUSERNAME+"&action=p";
//alert('调试模式');
if (pagenum==1){
 $.getJSON(targeturl,function(d){
    player(d.jsonObj[0]);
  });
}
}



function getweblist(url,domid,pagenum){
var CurrentFunPath;
var sUSERNAME;
var sGUID;
if(typeof(VideoPathRoot)=="undefined"){CurrentFunPath=""}else{CurrentFunPath=VideoPathRoot}
if(typeof(VideoUSERNAME)=="undefined"){sUSERNAME=""}else{sUSERNAME=VideoUSERNAME}
if(typeof(VideoGUID)=="undefined"){sGUID=""}else{sGUID=VideoGUID}
var objectn="#"+domid;
var VideoBaseURL;
if (CurrentFunPath=="") 
{VideoBaseURL=CurrentFunPath}
else
{
if(CurrentFunPath.indexOf('://')>=0)
{VideoBaseURL = CurrentFunPath;
}
else
{VideoBaseURL = 'http://'+window.location.host+'/';}
}
var targeturl=""+VideoBaseURL+"inc/ajaxtv.asp?callback=?&page="+pagenum+"&pagenum="+url+"&ntimest="+Math.random()+"&guid="+sGUID+"&u="+sUSERNAME;
 $.getJSON(targeturl,function(d){
    $(objectn).html(d.jsonObj[0]);autoplay(url,pagenum)
  });
}



function player(flvname){
  //播放器 , 播放器ID,宽 , 高,
  var s1 = new SWFObject("video/mediaplayer.swf","player_id",tvwidth,tvheight,"7");
  s1.addParam("allowfullscreen","true");//是否允许全屏播放
  s1.addVariable("file",flvname);//单文件播放
  s1.addVariable("image","/video/preview.jpg");//背景图片
  s1.addVariable("displayheight",tvheight);//播放区域高度
  s1.addVariable("width",tvwidth);
  s1.addVariable("height",tvheight);
  s1.addVariable("backcolor","0x000000");
  s1.addVariable("frontcolor","0xCCCCCC");
  s1.addVariable("lightcolor","0x557722");
  s1.addVariable("enablejs","true");//是否允许javascript脚本控制flash
  s1.addVariable("javascriptid","javascript_id");//控制脚本javascriptid
  s1.write("myPlayer");//将播放器写入到myPlayer
 }
function g(o){return document.getElementById(o);} 
function HoverLi(n){for(var i=1;i<=2;i++){g('tb_'+i).className='normaltab';g('tbc_0'+i).className='undis';}g('tbc_0'+n).className='dis';g('tb_'+n).className='hovertab';} 

//
//搜索ajax函数

function getajaxpage(url, keywords, perpagenum, ccid, action, pagenum) {
  var objectn = ".indexproductinner";
  var CurrentFunPath;
  var sUSERNAME;
  var sGUID;
  if (typeof(VideoPathRoot) == "undefined") {
    CurrentFunPath = ""
  } else {
    CurrentFunPath = VideoPathRoot
  }
  if (typeof(VideoUSERNAME) == "undefined") {
    sUSERNAME = ""
  } else {
    sUSERNAME = VideoUSERNAME
  }
  if (typeof(VideoGUID) == "undefined") {
    sGUID = ""
  } else {
    sGUID = VideoGUID
  }
  var VideoBaseURL;
  if (CurrentFunPath == "") {
    VideoBaseURL = CurrentFunPath
  } else {
    if (CurrentFunPath.indexOf('://') >= 0) {
      VideoBaseURL = CurrentFunPath;
    } else {
      VideoBaseURL = 'http://' + window.location.host + '/';
    }
  }
  var targeturl = ""
  if (url == "inc/taobao_search.asp") {
    targeturl = VideoBaseURL + url + "?callback=?&action=" + action + "&ccid=" + ccid + "&perpagenum=" + perpagenum + "&pagenum=" + pagenum + "&ntimest=" + Math.random() + "&guid=" + sGUID + "&u=" + sUSERNAME + "&userNick=" + userNick + "&keywords=" + keywords;
  } else {
    targeturl = VideoBaseURL + url + "?callback=?&action=" + action + "&ccid=" + ccid + "&perpagenum=" + perpagenum + "&pagenum=" + pagenum + "&ntimest=" + Math.random() + "&guid=" + sGUID + "&u=" + sUSERNAME + "&keywords=" + keywords;
  }
  $.getJSON(targeturl,function(d){
    $(objectn).html(d.jsonObj[0]);getajaxpagecomelete(objectn)
  });
}

//取得静态页参数函数
function getQueryString(queryStringName)
{
var returnValue="";
var URLString=new String(document.location);
var serachLocation=-1;
var queryStringLength=queryStringName.length;
do
{
   serachLocation=URLString.indexOf(queryStringName+"\=");
   if (serachLocation!=-1)
   {
    if ((URLString.charAt(serachLocation-1)=='?') || (URLString.charAt(serachLocation-1)=='&'))
    {
     URLString=URLString.substr(serachLocation);
     break;
    }
    URLString=URLString.substr(serachLocation+queryStringLength+1);
   }
  
}
while (serachLocation!=-1)
if (serachLocation!=-1)
{
   var seperatorLocation=URLString.indexOf("&");
   if (seperatorLocation==-1)
   {
    returnValue=URLString.substr(queryStringLength+1);
   }
   else
   {
    returnValue=URLString.substring(queryStringLength+1,seperatorLocation);
   } 
}
return returnValue;
}