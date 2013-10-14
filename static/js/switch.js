(function() {
    "use strict";
    $("#news_company").css('display', 'none');
    $("#news .switch>span").on('click', function(event) {
        //sdf
        if($(this).hasClass('act'))return;
        $("#news .switch>span.act").removeClass('act');
        $(this).addClass('act');
        if(this.innerHTML==="企业新闻"){
            $('#news_company').css('display','');
            $('#news_Industry').css('display','none');
        } else if(this.innerHTML === "行业新闻"){
            $('#news_company').css('display','none');
            $('#news_Industry').css('display','');
        }
    });
	$("#shows_product").css('display', 'none');
	$("#shows .switch>span").on('click', function(event) {
		//sdf
		if($(this).hasClass('act'))return;
		$("#shows .switch>span.act").removeClass('act');
		$(this).addClass('act');
		if(this.innerHTML==="产品展示"){
			$('#shows_case').css('display','none');
			$('#shows_product').css('display','');
		} else if(this.innerHTML === "工程案例"){
			$('#shows_case').css('display','');
			$('#shows_product').css('display','none');
		}
	});
})();