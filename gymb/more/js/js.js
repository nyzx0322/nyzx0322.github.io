$(function(){
	/*********幻灯*********/
	$('.news .content ul').width(300*$('.news .content li').length+'px');
	$(".news .tab a").mouseover(function(){
		$(this).addClass('on').siblings().removeClass('on');
		var index = $(this).index();
		number = index;
		var distance = -300*index;
		$('.news .content ul').stop().animate({
			left:distance
		},150);
	});
	/**
	var auto = 1;  //等于1则自动切换，其他任意数字则不自动切换
	if(auto ==1){
		var number = 0;
		var maxNumber = $('.news .tab a').length;
		function autotab(){
			number++;
			number == maxNumber? number = 0 : number;
			$('.news .tab a:eq('+number+')').addClass('on').siblings().removeClass('on');
			var distance = -300*number;
			$('.news .content ul').stop().animate({
				left:distance
			});
		}
		var tabChange = setInterval(autotab,3000);
		//鼠标悬停暂停切换
		$('.news').mouseover(function(){
			clearInterval(tabChange);
		});
		$('.news').mouseout(function(){
			tabChange = setInterval(autotab,3000);
		});
	  }  
	  **/
	/*********hover*********/
	$(".web_list>li>div").css("background-color","#F66").hide();
	$(".web_list>li").mouseover(function(){
		$(this).children("div").stop().fadeIn(100);
	});
	$(".web_list>li").mouseout(function(){
		$(this).children("div").stop().fadeOut(200);
	});
	/*********向上*********/
	$("#back").click(function(){
		$('body,html').animate({scrollTop:0},1000,'easeOutQuint');
		return false;
	});
});
