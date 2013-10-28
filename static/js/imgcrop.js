var $img = $('li img'); //获取图片
$img.each(function() {
    var newImg = new Image(); //建立新的图片对象
    newImg.src = $(this).attr('src');
    var pic_real_height = newImg.height;
    var pic_real_width = newImg.width;
    //alert(pic_real_height + ',' + pic_real_width); //打印检测
});
var ImgCrops;
(function() {
    "use strict";
    ImgCrops = function(node) {
        $(node + ' img').load(function() {
            var targetWidth = $(this).css('width');
            var targetHeight = $(this).css('height');
            this.style.width = this.style.height = 'auto';
            //var newImg = new Image();
            //newImg.src = $(this).attr('src');
            //var pic_real_height = newImg.height;
            //var pic_real_width = newImg.width;
            ////alert(newImg.height)
            if(parseInt($(this).css('width')) > parseInt($(this).css('height'))) {
                //console.log('asdf');
                this.style.height = targetHeight;
                this.style.marginLeft = (parseInt(targetWidth) - this.offsetWidth) / 2 + 'px';
                //alert(this.offsetWidth);
            } else if(parseInt($(this).css('width')) < parseInt($(this).css('height'))) {
                //console.log('odibt');
                this.style.width = targetWidth;
                this.style.marginTop = (parseInt(targetHeight) - this.offsetHeight) / 2 + 'px';
            } else {
                this.style.width = targetWidth;
            }
            //var va = Math.min($(this).css('width'), $(this).css('height'))
            var div = this.parentNode.insertBefore(document.createElement("DIV"), this);
            div.style.overflow = 'hidden';
            div.style.margin = $(this).css('padding');
            div.style.width = targetWidth;
            div.style.height = targetHeight;
            div.appendChild(this);
            //alert(this.clientWidth)
            $(this).css('padding', 0);
            //this.width = pic_real_width;
            //this.height = pic_real_height;
            //console.log('pic_real_height:'+pic_real_height);
        });
    };
})();
