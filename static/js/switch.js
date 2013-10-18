var bindSwitch;
(function() {
    "use strict";
    bindSwitch = function(target, Obj) {
        /*Obj = {'企业新闻':'news_company','行业新闻':'news_Industry'}*/
        $("#" + target + " .switch>span").each(function() {
            if($(this).hasClass('act')) {
                $("#" + Obj[this.innerHTML]).css('display', '');
            } else {
                $("#" + Obj[this.innerHTML]).css('display', 'none');
            }
        });
        $("#" + target + " .switch>span").on('click', function(event) {
            //sdf
            if($(this).hasClass('act'))return;
            $("#" + target + " .switch>span.act").removeClass('act');
            $(this).addClass('act');
            for(var iii in Obj) {
                $("#"+Obj[iii]).css('display','none');
            }
            $("#"+Obj[this.innerHTML]).css('display', '');
        });
    };
})();