//**********************************************************
//* Relative Positioning Module for DHTML Menu Builder 2.x *
//* Coded by Xavier Flix                                   *
//* ------------------------------------------------------ *
//* Version 0.1                                            *
//**********************************************************


	//CUSTOM VARIABLES
	var NavBarWidth = 600;
	var NavBarHeight = 20;
	var NavItemsWidth = 40;
	
	//SYSTEM VARIABLES & FUNCTIONS
	var cScroll = 0;

	function GetItemLeftPos(itm) {
		if(IE)
			return parseInt(NavBar.style.left) + parseInt(itm.style.left) + "px";
		if(NS)
			return document.layers.NavBar.left + itm.left + "px";
	}
	
	function GetItemTopPos(itm) {
		if(IE)
			return parseInt(NavBar.style.top) + parseInt(itm.style.top) + parseInt(itm.style.height) + "px";
		if(NS)
			return document.layers.NavBar.top + itm.top + itm.height + "px";
	}

	function CenterNavBar() {
		if(IE) {
			var cX = document.body.clientWidth;
			var cY = document.body.clientHeight;
			var n = NavBar.style;
		}
		
		if(NS) {
			var cX = window.innerWidth;
			var cY = window.innerHeight;
			var n = document.layers.NavBar;
			n.resizeTo(NavBarWidth, NavBarHeight);
			n.moveTo(0,window.pageYOffset);
			for(var i=0; i<=n.layers.length-1; i++) {
				n.layers[i].height = NavBarHeight;
				n.layers[i].width = NavItemsWidth;
			}
		}
	
		n.width = NavBarWidth;
		n.height = NavBarHeight;
		n.left = (cX - parseInt(n.width))/2;
		if(IE) n.top = document.body.scrollTop;
		if(NS) n.top = window.pageYOffset;
		n.visibility = "visible";
		
		if(NS)
			window.setTimeout(MonitorScroll,100);
	}
	
	function MonitorScroll() {
		if(cScroll!=window.pageYOffset) {
			CenterNavBar();
			cScroll = window.pageYOffset;
		}
		window.setTimeout(MonitorScroll,100);
	}
	
	window.onload = CenterNavBar;
	window.onresize = CenterNavBar;
	window.onscroll = CenterNavBar;
