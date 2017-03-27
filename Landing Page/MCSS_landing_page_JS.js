
//jQuery to collapse the navbar on scroll -> scrolling Nav
$(window).scroll(function() {
    if ($(".navbar").offset().top > 50) {
        $(".navbar-fixed-top").addClass("top-nav-collapse");
    } else {
        $(".navbar-fixed-top").removeClass("top-nav-collapse");
    }
});

//jQuery for page scrolling feature - requires jQuery Easing plugin
$(function() {
    $('a.page-scroll').bind('click', function(event) {
        var $anchor = $(this);
        $('html, body').stop().animate({
            scrollTop: $($anchor.attr('href')).offset().top
        }, 1500, 'easeInOutExpo');
        event.preventDefault();
    });
});

//full calendar

var cal_api_password = 'AIzaSyAKJ7Q-SRhyq9CZBHyw6aoYnwYXfEqXbhs';

$(document).ready(function() {
    $('#calendar').fullCalendar({
        googleCalendarApiKey: cal_api_password,
        eventSources: [
            {
                googleCalendarId: 'pardotadvocates@gmail.com',
                className: 'OOH',
                signup: 'Please login at the event time:',
                text: 'Int. dialing codes found: <a class="read-more" href="http://www2.pardot.com/advocatesopenoffice" target="_blank">here</a>',
                team: 'Client Advocates'
            },
            {
                googleCalendarId: 'smh76g1ak9ht58n9gtv4fjolis@group.calendar.google.com',
                className: 'TOH',
                signup: 'Please register:',
                text: 'dialing info will be emailed after event registration',
                color: '#9adae0',
                textColor: 'black',
                team: 'Client Advocates'
            },
            {
                googleCalendarId: '03g6lb5m6hb697vsbpgt0ov5d4@group.calendar.google.com',
                className: 'newuser',
                signup: 'Please register:',
                text: 'dialing info will be emailed after event registration',
                color: '#ffad46',
                textColor: 'black',
                team: 'Client Advocates'
            },
            {
                googleCalendarId: '0a5evir5adnhu540jdrs0763lg@group.calendar.google.com',
                className: 'training',
                signup: 'Please register:',
                text: 'dialing info will be emailed after event registration',
                color: '#b3dc6c',
                textColor: 'black',
                team: 'Pardot Training & Certification'
            },
            {
	            googleCalendarId: 'e9movuvl0obl4b4a5s107dplq0@group.calendar.google.com',
	            className: 'cert',
	            signup: 'Please register:',
	            text: 'dialing info will be emailed after event registration. <br>More information on Pardot certification can be found <a class="read-more" href="http://www.pardot.com/user-certification" target="_blank">here</a></br>',
	            color: '#f691b2',
                textColor: 'black',
                team: 'Pardot Training & Certification'
            },
            {
	            googleCalendarId: '2johr6ok3gpr3f6tdjifaebnk0@group.calendar.google.com',
	            text: 'dialing info will be emailed after event registration',
	            signup: 'Please register:',
	            className: 'webinars',
	            color: '#b767ce',
	            team: 'Pardot Marketing'
            },
            {
	            googleCalendarId: 'i5u7lkkpc6o7qqruon6i2dcckg@group.calendar.google.com',
	            text: 'Please only register if you will be attending DreamForce / World Tour conference',
	            signup: 'Please register:',
	            className: 'events',
	            color: '#4ea96c',
	            team: 'Pardot Marketing'
            },
            {
	            googleCalendarId: '35fipfulhgr3r62kib3itej4v8@group.calendar.google.com',
	            text: 'Please register to attend a user group in person',
	            signup: 'Please register:',
	            className: 'usergroup',
	            color: '#fffd09',
                textColor: 'black',
                team: 'Pardot Marketing'
            }
        ],
        eventRender: function (event, element) {
        	element.attr('href', 'javascript:void(0);');
            element.qtip({
	            style: {
		            classes: 'qtip-light qtip-rounded qtip-shadow'
		            },
		        position: {
					my: 'bottom left',
        	        at: 'top right',
                    viewport: $(window),
					adjust: {
						mouse: false,
						scroll: false,
                        method: 'shift flip',
                        resize: true
							}
    				},
                show: {when:'click', solo: true},
				hide: 'unfocus',
	            content: {
		            title: event.title,
		            text: 
		            '<p id="qtip">'+event.description+'</p>'
		            +'<p id="qtip">' +event.source.signup+' <a class="read-more" href="'+event.location+current_time(event.start, event.end)+'" target="_blank">here</a></p>'
		            +'<p id="qtip">'+event.source.text+'</p>'
		            +'<p id="qtip"><b>Delivered by: </b>'+event.source.team+'</p>'
		            +'<p id="qtip"><b>Start: </b>'+moment(event.start).format("DD MMM YY - HH:mm")+'</p>'
		            +'<p id="qtip"><b>End: </b>'+moment(event.end).format("DD MMM YY - HH:mm")+'</p>'
		            },
	            });
    },
    	viewRender:function(){
	    	$('#sessionform input:checkbox').each(function(){
              $(this).prop('checked', true);
            });},
    	businessHours: {start: "08:00:00", end: "18:00:00"},
        contentHeight: 'auto',
         //scrollTime: "06:00:00",
        minTime: "06:00:00",
        maxTime: "22:00:00",
       	nowIndicator: true,
        weekends: false,
        timezone: 'local',
        allDaySlot: false,
	    defaultView: 'agendaWeek',
	    header: {
				left: 'prev,next today',
				center: 'title',
				right: 'agendaWeek,month'
			}
    });


//calendar category toggle

$('#sessionform input:checkbox').on('change',function(){
	var name = $(this).attr('name');
	if (this.checked)
		{$("."+name).css({'visibility': 'visible'});}
	if (!this.checked)
		{$("."+name).css({'visibility': 'hidden'});}
});

//slide toggle calendar filters

$('#slidetog').on('click', function(){$("#slide").slideToggle();});


});

//is event happening now parameter

function current_time(event_start, event_end) {
	var nowtime = moment($('#calendar').fullCalendar('getDate')).format();
	var eventstart = moment(event_start).subtract(5, 'm').format();
	var eventend = moment(event_end).subtract(5, 'm').format();
	if(moment(nowtime).isBetween(eventstart, eventend))
		{return "?ooh=y";}
	else 
		{return "";};
	};

//category sticky to side of screen

function sticky_relocate() {
    var window_top = $(window).scrollTop();
    var div_top = $('#sticky-anchor').offset().top;
    var end = $('#calendar_stop').offset().top;
    	
    if (window_top < div_top) {
	    $('#sticky').removeClass('stick');
	    $('#slidetog').hide();
	    $('#slide').show();	    
    }
    else if (window_top > end) {
        $('#sticky').removeClass('stick');
        $('#slide').hide();
        $('#slidetog').hide();
    }
    else if (window_top > div_top) {
	    $('#sticky').addClass('stick');
        $('#slide').hide();	
        $('#slidetog').show();	 
    };
};

$(function () {
    $(window).scroll(sticky_relocate);
    sticky_relocate();
});

//local timezone notification 
	
$(function timezone() {
	var tz = jstz.determine();
	var timename = tz.name();
	var timezone = moment.tz.zone(tz.name()).abbr(new Date().getTime());
	$("#tmz_name").append(timename);
    $("#tmz_zone").append(timezone);
});

//fancybox and wistia video player

    $(document).ready(function() {
     $(".fancybox").fancybox({
        'type': 'iframe',
        //'fitToView' : true,
		'autoSize'  : true,
		'width'    : 1000,
		//'height'   : 800,
		'openEffect' : false,
		'closeEffect': false,
		helpers: {
			overlay: {
				locked: false
    }
  }
    });
});

var count = 0;
var shown_videos = 12; 
var playlist_ID = "kz5clc3mub";
var video_api_password = "9c3639e2ccc71cab740aff449ab77047d6f7ea7fa77097f0fac5b862f4fe5eb8";
$.getJSON('https://api.wistia.com/v1/medias.json?api_password='+video_api_password+'&sort_by=created&sort_direction=0', function(data) {
    $.each(data, function(index, video) {
	    if(video.project.hashed_id == playlist_ID)
	    	{count++;
			 if(count < shown_videos)
			 	{
			 	$('#player').append(
				 	'<div class="video_box col-xs-12 col-sm-6 col-md-3 col-lg-3">'+
				 			'<div class="video_name">'+video.name+'</div>'+
				 			'<a class="video_thumb fancybox fancybox.iframe" href="https://fast.wistia.net/embed/iframe/'+video.hashed_id+'">'+
				 				'<img src='+video.thumbnail.url.split("?")[0]+'>'+
				 			'</a>'+
				 			'<i class="fa fa-play-circle"></i>'+
				 	'</div>'	
			 		);
			 	}
			 else
			 	{return false;}}
})});    

//success community web scraper

$.getJSON('http://allorigins.pw/get?url=' + encodeURIComponent('https://success.salesforce.com/featuredGroupDetail?id=a1z30000006IDZ5AAO') + '&callback=?', function(data){
        var feed = $(data.contents).find(".list-of-items");
        $("#feed_output").append(feed);
        
        var members = $(data.contents).find(".topic-title b").text();
        $("#success_sub").append(members);
        
        var topics = $(data.contents).find(".chatter-topics-content p").text();
        $("#topic_output").append(topics);
});

//load images with salesforce domain and after web scrap

$(window).bind("load",function(){
$('div.user-pic img').each(function(){
    $(this).attr('src', 'https://success.salesforce.com' + $(this).attr('src'));
})});