var magnifier;
(function() {
    "use strict";
    magnifier = function(id) {
        $("#" + id + " img").mouseover(function(event) {
            if(!document.getElementById('magnifier')) {
                (document.body.appendChild(document.createElement("div")).id = 'magnifier');
                $('#magnifier').css({
                    position: 'absolute',
                    width: 316,
                    height: 308,
                    background: 'url(template/sweelife/image/bg_product_big.png)',
                    zIndex: 1000
                })
                var img = document.getElementById('magnifier').appendChild(document.createElement('IMG'));
                $(img).css({
                    width:277,
                    height:270,
                    margin:'17px 1px 1px 18px'
                })
            }

            $('#magnifier').css({
                display: '',
                left: ((document.documentElement.clientWidth - event.pageX - 316 - 30) < 0) ? event.pageX - 20 - 316 : event.pageX + 20,
                top: event.pageY - 80
            })
            $('#magnifier img').attr('src', this.src)
            console.log(((document.documentElement.clientWidth - event.pageX - 316) ));

        }).mouseleave(function() {
                $('#magnifier').css('display', 'none');
            })
    };
})();