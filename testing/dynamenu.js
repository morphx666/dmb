////////////////////////////////////////////////
//  DHTML MENU BUILDER 2.6.161                //
//  (c)xFX JumpStart                          //
//                                            //
//  PSN: 7294-238782-XFX-4540                 //
//                                            //
//  GENERATED: 8/20/2000 - 2:13:19 AM         //
////////////////////////////////////////////////


	var nStyle = new Array;
	var hStyle = new Array;
	var nLayer = new Array;
	var hLayer = new Array;
	var nTCode = new Array;

	var AnimStep = 0;
	var AnimHnd = 0;
	var HTHnd = new Array;
	var DoFormsTweak = false;

	var mFrame;
	var cFrame;

	var OpenMenus = new Array;
	var SelCommand;
	var nOM = 0;

	var mX;
	var mY;

	var HideSpeed = 200;

var BV=parseInt(navigator.appVersion);
var BN=window.navigator.appName;
var IsMac=(navigator.userAgent.indexOf('Mac')!=-1)?true:false;
var Opera=(navigator.userAgent.indexOf('Opera')!=-1)?true:false;
var NS=(BN.indexOf('Netscape')!=-1&&(BV==4)&&!Opera)?true:false;
var IE=(BN.indexOf('Explorer')!=-1&&(BV>=4)&&!Opera)?true:false;


	if ((frames.length==0) && IsMac) {
		frames.top = window;
	}
	cFrame = eval(frames['top']);

	var fx = 0;



	nStyle[0]="border: 1px;  font-family: Tahoma; font-size: 9pt; font-weight: none; font-style: none; color: #000000; background-color: #C0C0C0; cursor: hand";
	hStyle[0]="border: 1px;  font-family: Tahoma; font-size: 9pt; font-weight: none; font-style: none; color: #FFFFFF; background-color: #0000FF; cursor: hand;";
	nStyle[1]="border: 1px;  font-family: Tahoma; font-size: 9pt; font-weight: none; font-style: none; color: #000000; background-color: #C0C0C0; cursor: hand";
	hStyle[1]="border: 1px;  font-family: Tahoma; font-size: 9pt; font-weight: none; font-style: none; color: #FFFFFF; background-color: #0000FF; cursor: hand;";
	nStyle[2]="border: 1px;  font-family: Tahoma; font-size: 9pt; font-weight: none; font-style: none; color: #000000; background-color: #C0C0C0; cursor: hand";
	hStyle[2]="border: 1px;  font-family: Tahoma; font-size: 9pt; font-weight: none; font-style: none; color: #FFFFFF; background-color: #0000FF; cursor: hand;";
	nTCode[1]="frames[\'top\'].execURL(\'\', \'frames[top]\');";
	nLayer[1]="<p align=left><font face=Tahoma point-size=9 color=#000000>Menu 1 Option 1</font>";
	hLayer[1]="<p align=left><font face=Tahoma point-size=9 color=#FFFFFF>Menu 1 Option 1</font>";
	nTCode[2]="frames[\'top\'].execURL(\'\', \'frames[top]\');";
	nLayer[2]="<p align=left><font face=Tahoma point-size=9 color=#000000>Menu 1 Option 2</font>";
	hLayer[2]="<p align=left><font face=Tahoma point-size=9 color=#FFFFFF>Menu 1 Option 2</font>";
	nTCode[3]="frames[\'top\'].execURL(\'\', \'frames[top]\');";
	nLayer[3]="<p align=left><font face=Tahoma point-size=9 color=#000000>Dynamic Menu</font>";
	hLayer[3]="<p align=left><font face=Tahoma point-size=9 color=#FFFFFF>Dynamic Menu</font>";


	function GetCurCmd() {
//IE
//This function will return the current command under the mouse pointer.
//It will return null if the mouse is not over any command.
//------------------------------
//Version 1.3
//
		var cc = mFrame.window.event.srcElement;
		while((cc.id=="") && (cc.tagName!="TD")) {
			cc = cc.parentElement;
			if(cc==null)
				break;
		}
		return cc;
	}

	function HoverSel(mode, imgLName, imgRName) {
//IE
//This is the function called every time the mouse pointer is moved over a command.
//------------------------------
//mode: 0 if the mouse is moving over the command and 1 if is moving away.
//imgLName: Name of the left image object, if any.
//imgRName: Name of the right image object, if any.
//------------------------------
//Version 9.1
//
		var imgL = new Image;
		var imgLRsc = new Image;
		var imgR = new Image;
		var imgRRsc = new Image;
		var mc;

		if(mode==0 && OpenMenus[nOM].SelCommand!=null)
			HoverSel(1);

		if(imgLName!="_")
			var imgL = eval("mFrame.document.images['"+imgLName+"']");
		if(imgRName!="_")
			var imgR = eval("mFrame.document.images['"+imgRName+"']");

		if(mode==0) {
			mc = GetCurCmd();
			if(nOM>1)
				if(mc==OpenMenus[nOM-1].SelCommand)
					return;
			while(!InMenu())
				Hide();
			mc.style.cssText = hStyle[mc.id];
			if(imgLName!='_') imgLRsc = eval(imgLName+"On");
			if(imgRName!='_') imgRRsc = eval(imgRName+"On");
			OpenMenus[nOM].SelCommand = mc;
			OpenMenus[nOM].SelCommandPar = [imgLName,imgRName];
		} else {
			mc = (mode==1)?OpenMenus[nOM].SelCommand:OpenMenus[nOM].Opener;
			imgLName = (mode==1)?OpenMenus[nOM].SelCommandPar[0]:OpenMenus[nOM].OpenerPar[0];
			imgRName = (mode==1)?OpenMenus[nOM].SelCommandPar[1]:OpenMenus[nOM].OpenerPar[1];
			mc.style.background = "";
			mc.style.cssText = nStyle[mc.id];
			if(imgLName!='_') imgLRsc = eval(imgLName+"Off");
			if(imgRName!='_') imgRRsc = eval(imgRName+"Off");
			window.status = "";
			OpenMenus[nOM].SelCommand = null;			
		}

		if(imgLName!='_') {
			imgL = eval("mFrame.document.images."+imgLName);
			imgL.src = imgLRsc.src;
		}
		if(imgRName!='_') {
			imgR = eval("mFrame.document.images."+imgRName);
			imgR.src = imgRRsc.src;
		}
	}

	function NSHoverSel(mode, mc, bcolor, w, h) {
//NS
//This is the function called every time the mouse pointer is moved over or away from a command.
//------------------------------
//mode: 0 if the mouse is moving over the command and 1 if is moving away.
//mc: Name of the layer that corresponds to the selected command.
//n: Unique ID that identifies this command. Used to retrieve the data from the nLayer or hLayer array.
//bcolor: Background color of the command. Ignored if the group uses a background image.
//w: Width of the command's layer.
//h: Height of the command's layer.
//------------------------------
//Version 10.3
//
		var n;
		var LayerHTM;
		if(mode==0 && OpenMenus[nOM].SelCommand!=null)
			NSHoverSel(1);
		
		if(mode==0) {
			n = parseInt(mc.name.substr(2));
			if(nOM>1)
				if(mc==OpenMenus[nOM-1].SelCommand)
					return;
			while(!InMenu())
				Hide();
			LayerHTM = hLayer[n];
			OpenMenus[nOM].SelCommand = mc;
			OpenMenus[nOM].SelCommandPar = [mc.bgColor,w,h];
			mc.bgColor = bcolor;
		} else {
			mc = (mode==1)?OpenMenus[nOM].SelCommand:OpenMenus[nOM].Opener;
			bcolor = (mode==1)?OpenMenus[nOM].SelCommandPar[0]:OpenMenus[nOM].OpenerPar[0];
			w = (mode==1)?OpenMenus[nOM].SelCommandPar[1]:OpenMenus[nOM].OpenerPar[1];
			h = (mode==1)?OpenMenus[nOM].SelCommandPar[2]:OpenMenus[nOM].OpenerPar[2];
			n = parseInt(mc.name.substr(2));
			LayerHTM = nLayer[n];
			if(mc.parentLayer.background.src!="")
				mc.bgColor = null;
			else
				mc.bgColor = bcolor;
			window.status = "";
			OpenMenus[nOM].SelCommand = null;
		}
		mc.resizeTo(w,h);
		mc.document.open();
		mc.document.write(LayerHTM);
		mc.document.close();
	}

	function Hide() {
//IE,NS
//This function hides the last opened group and it keeps hiding all the groups until
//no more groups are opened or the mouse is over one of them.
//Also takes care of reseting any highlighted commands.
//------------------------------
//Version 3.1
//

		if(AnimHnd)
			window.clearTimeout(AnimHnd);

		if(OpenMenus[nOM].SelCommand!=null) {
			if(IE) HoverSel(1);
			if(NS) NSHoverSel(1);
		}
		if(OpenMenus[nOM].Opener!=null) {
			if(IE) HoverSel(3);
			if(NS) NSHoverSel(3);
		}

		OpenMenus[nOM].visibility = "hidden";
		window.clearTimeout(HTHnd[nOM]);
		HTHnd[nOM] = 0;
		nOM--;

		if(nOM>0)
			if(!InMenu())
				HTHnd[nOM] = window.setTimeout("Hide()", HideSpeed);
	}

	function ShowMenu(mName, x, y, isCascading) {
//IE,NS
//This is the main function to show the menus when a hotspot is triggered or a cascading command is activated.
//------------------------------
//mName: Name of the <div> or <layer> to be shown.
//x: Left position of the menu.
//y: Top position of the menu.
//isCascading: True if the menu has been triggered from a command, and not from a hotspot.
//------------------------------
//Version 13.7
//
		x = parseInt(x);
		y = parseInt(y);
		if(AnimHnd && nOM>0) {
			AnimStep=100;
			Animate();
		}
		if(IE)
			var Menu = mFrame.document.all[mName].style;
		if(NS)
			var Menu = mFrame.document.layers[mName];
		if(Menu==OpenMenus[nOM] || HTHnd[nOM])
			return;
		
		Menu.Opener = nOM>0?OpenMenus[nOM].SelCommand:null;
		Menu.OpenerPar = nOM>0?OpenMenus[nOM].SelCommandPar:null;
		Menu.SelCommand = null;
		
		if(!isCascading)
			HideAll();

		var pW = GetWidthHeight()[0] + GetLeftTop()[0];
		var pH = GetWidthHeight()[1] + GetLeftTop()[1];
		
		if(IE) {
			if(isCascading) {
				x = parseInt(OpenMenus[nOM].left) + parseInt(OpenMenus[nOM].width) - 6;
				y = y + parseInt(OpenMenus[nOM].top) - 5;
				Menu.left = (x+parseInt(Menu.width)>pW)?parseInt(OpenMenus[nOM].left)-parseInt(Menu.width) + 6:x;
				Menu.top =  (y+parseInt(Menu.height)>pH)?pH-parseInt(Menu.height):y;
			} else {
				Menu.left = (x+parseInt(Menu.width)>pW)?pW-parseInt(Menu.width):x;
				Menu.top =  (y+parseInt(Menu.height)>pH)?pH-parseInt(Menu.height):y;
			}
			if(!IsMac)
				Menu.clip = "rect(0 0 0 0)";
		}
		if(NS) {
			if(isCascading) {
				x = OpenMenus[nOM].left + OpenMenus[nOM].clip.width - 6;
				y = OpenMenus[nOM].top + OpenMenus[nOM].SelCommand.top;
				x = (x+Menu.w>pW)?OpenMenus[nOM].left-Menu.w + 6:x;
				y = (y+Menu.h>pH)?pH-Menu.h:y;
			} else {
				x = (x+Menu.w>pW)?pW-Menu.w:x;
				y = (y+Menu.h>pH)?pH-Menu.h:y;
			}
			Menu.clip.width=0;
			Menu.clip.height=0;
			Menu.moveToAbsolute(x,y);
		}
		if(isCascading)
			Menu.zIndex = parseInt(OpenMenus[nOM].zIndex) + 1;
		Menu.visibility = "visible";
		OpenMenus[++nOM] = Menu;
		HTHnd[nOM] = 0;
		if(!IsMac)
			AnimHnd = window.setTimeout("Animate()", 10);
		FormsTweak("hidden");
	}

	function Animate() {
//IE,NS
//This function is called by ShowMenu every time a new group must be displayed and produces the predefined unfolding effect.
//Currently is disabled for Navigator, because of some weird bugs we found with the clip property of the layers.
//------------------------------
//Version 1.7
//
		var r = '';
		var nw = nh = 0;
		switch(fx) {
			case 1:
				if(IE) r = "0 " + AnimStep + "% " + AnimStep + "% 0";
				if(NS) nw = AnimStep; nh = AnimStep;
				break;
			case 2:
				if(IE) r = "0 100% " + AnimStep + "% 0";
				if(NS) nw = 100; nh = AnimStep;
				break;
			case 3:
				if(IE) r = "0 " + AnimStep + "% 100% 0";
				if(NS) nw = AnimStep; nh = 100;
				break;
			case 0:
				if(IE) r = "0 100% 100% 0";
				if(NS) nw = 100; nh = 100;
				break;
		}
		with(OpenMenus[nOM]) {
			if(IE)
				clip =  "rect(" + r + ")";
			if(NS) {
				clip.width = w*(nw/100);
				clip.height = h*(nh/100);
			}
		}
		AnimStep += 20;
		if(AnimStep<=100)
			AnimHnd = window.setTimeout("Animate()",25);
		else {
			window.clearTimeout(AnimHnd);
			AnimStep = 0;
			AnimHnd = 0;
		}
	}
	
	function InMenu() {
//IE,NS
//This function returns true if the mouse pointer is over the last opened menu.
//------------------------------
//Version 1.4
//
		var m = OpenMenus[nOM];
		var l = parseInt(m.left) + 2;
		var r = l+((IE)?parseInt(m.width):m.clip.width) - 2;
		var t = parseInt(m.top) + 2;
		var b = t+((IE)?parseInt(m.height):m.clip.height) - 2;
		return ((mX>=l && mX<=r) && (mY>=t && mY<=b));
	}

	function HideMenus(e) {
//IE,NS
//This function checks if the mouse pointer is on a valid position and if the current menu should be kept visible.
//The function is called every time the mouse pointer is moved over the document area.
//------------------------------
//e: Only used under Navigator, corresponds to the Event object.
//------------------------------
//Version 24.0
//
		if(IE) {
			if(event==null)
				e = mFrame.window.event;
			else
				e = event;
			mX = e.clientX + mFrame.document.body.scrollLeft;
			mY = e.clientY + mFrame.document.body.scrollTop;
		}
		if(NS) {
			mX = e.pageX + window.pageXOffset;
			mY = e.pageY + window.pageYOffset;
		}
		
		if(nOM>0)
			if(OpenMenus[nOM].SelCommand!=null)
				while(!InMenu() && !HTHnd[nOM]) {
					HTHnd[nOM] = window.setTimeout("Hide()", HideSpeed);
					if(nOM==0)
						break;
				}
	}
	
	function FormsTweak(str) {
//IE
//This is an undocumented function, which can be used to hide every form element on a page.
//This can be useful if the menus will be displayed over an area where is a combo box, which is an element that cannot be placed behind the menus and it will always appear over the menus resulting in a very undesirable effect.
//------------------------------
//Version 1.0
//
		if(DoFormsTweak)
			for(var i = 0; i <= (mFrame.document.forms.length - 1); i++)
				mFrame.document.forms[i].style.visibility = str;
	}

	function execURL(url, tframe) {
//IE,NS
//This function is called every time a command is triggered to jump to another page or execute some javascript code.
//------------------------------
//url: Encrypted URL that must be opened or executed.
//tframe: If the url is a document location, tframe is the target frame where this document will be opened.
//------------------------------
//Version 1.1
//
		HideAll();
		window.setTimeout("execURL2('" + url + "', '" + tframe + "')", 100);
	}

	function execURL2(url, tframe) {
//IE,NS
//This function is called every time a command is triggered to jump to another page or execute some javascript code.
//------------------------------
//url: Encrypted URL that must be opened or executed.
//tframe: If the url is a document location, tframe is the target frame where this document will be opened.
//------------------------------
//Version 1.0
//
		tframe = rStr(tframe);
		var fObj = eval(tframe);
		url = rStr(url);
		if(url.indexOf("javascript")!=url.indexOf("vbscript"))
			eval(url);
		else
			fObj.location.href = url;
	}

	function rStr(s) {
//IE,NS
//This function is used to decrypt the URL parameter from the triggered command.
//------------------------------
//Version 1.1
//
		s = xrep(s,"\x1E","'");
		s = xrep(s,"\x1D","\x22");
		s = xrep(s,"\x1C",",");
		return s;
	}

	function xrep(s, f, n) {
//IE,NS
//This function looks for any occurrence of the f string and replaces it with the n string.
//------------------------------
//Version 1.0
//
		var tmp = s.split(f);
		return tmp.join(n);
	}

	function hNSCClick(e) {
//NS
//This function executes the selected command's trigger code.
//------------------------------
//Version 1.0
//
		eval(this.TCode);
	}

	function HideAll() {
//IE,NS
//This function will hide all the currently opened menus.
//------------------------------
//Version 1.0
//
		while(nOM>0)
			Hide();
	}

	function GetLeftTop() {
//IE,NS
//This function returns the scroll bars position on the menus frame.
//------------------------------
//Version 1.0
//
		if(IE)
			return [mFrame.document.body.scrollLeft,mFrame.document.body.scrollTop];
		if(NS)
			return [mFrame.pageXOffset,mFrame.pageYOffset];
	}

	function GetWidthHeight() {
//IE,NS
//This function returns the width and height of the menus frame.
//------------------------------
//Version 1.0
//
		if(IE)
			return [mFrame.document.body.clientWidth,mFrame.document.body.clientHeight];
		if(NS)
			return [mFrame.innerWidth,mFrame.innerHeight];
	}

	function SetUpEvents() {
//IE,NS
//This function initializes the frame variables and setups the event handling.
//------------------------------
//Version 1.0
//
		if(typeof(mFrame)=="undefined")
		mFrame = eval(frames['top']);
		if(typeof(mFrame)=="undefined")
			window.setTimeout("SetUpEvents()",10);
		else {
			if(NS) {
				mFrame.captureEvents(Event.MOUSEMOVE);
				mFrame.onmousemove = HideMenus;
				PrepareEvents();
			}
			mFrame.document.onmousemove = HideMenus;
			document.onmousemove = HideMenus;
		}
	}

	function PrepareEvents() {
//NS
//This function is called right after the menus are rendered.
//It has been designed to attach the OnClick event to the <layer> tag. This is being
//done this way because Navigator does not support a click inline event capturing on
//the layer tag... duh!
//------------------------------
//Version 2.0
//
		for(var l=0; l<mFrame.document.layers.length; l++) {
			var lo = mFrame.document.layers[l];
			lo.w = lo.clip.width;
			lo.h = lo.clip.height;
			for(var sl=0; sl<lo.layers.length; sl++) {
				var slo = mFrame.document.layers[l].layers[sl];
				if(slo.name.indexOf("EH")>0) {
					slo.document.captureEvents(Event.CLICK);
					slo.document.onclick = hNSCClick;
					slo.document.TCode = nTCode[slo.name.split("EH")[1]];
				}					
			}
		}
	}

	if(IE)
		with(document) {
			open();
			write("<div id=\"menu1\" style=\"position: absolute; top: 0%; left: 0%; width: 111; height: 44; z-index: 100; visibility: hidden;\"><table id=\"dmbMenu\" background=\"\" border=\"0\" cellpadding=\"1\" style=\"background-color: #C0C0C0; color: #000000; border-left: #E0E0E0 solid 1; border-right: 1 solid #808080; border-top: 1 solid #E0E0E0; border-bottom: 1 solid #808080\" width=\"111\"><tr><td nowrap height=\"18\" align=\"left\" style=\"border: 1px; font-family: Tahoma; font-size: 9pt; font-weight: none; font-style: none; color: #000000; cursor: hand; background-color: #C0C0C0;\" id=\"0\" OnMouseOver=\"frames[\'top\'].HoverSel(0,\'_\',\'_\');window.status=\'Menu 1 Option 1\';\" OnClick=\"frames[\'top\'].execURL(\'\', \'frames[top]\');\">Menu 1 Option 1</td></tr><tr><td nowrap height=\"18\" align=\"left\" style=\"border: 1px; font-family: Tahoma; font-size: 9pt; font-weight: none; font-style: none; color: #000000; cursor: hand; background-color: #C0C0C0;\" id=\"1\" OnMouseOver=\"frames[\'top\'].HoverSel(0,\'_\',\'_\');window.status=\'\';\" OnClick=\"frames[\'top\'].execURL(\'\', \'frames[top]\');\">Menu 1 Option 2</td></tr></table></div><div id=\"menu2\" style=\"position: absolute; top: 0%; left: 0%; width: 97; height: 24; z-index: 100; visibility: hidden;\"><table id=\"dmbMenu\" background=\"\" border=\"0\" cellpadding=\"1\" style=\"background-color: #C0C0C0; color: #000000; border-left: #E0E0E0 solid 1; border-right: 1 solid #808080; border-top: 1 solid #E0E0E0; border-bottom: 1 solid #808080\" width=\"97\"><tr><td nowrap height=\"18\" align=\"left\" style=\"border: 1px; font-family: Tahoma; font-size: 9pt; font-weight: none; font-style: none; color: #000000; cursor: hand; background-color: #C0C0C0;\" id=\"2\" OnMouseOver=\"frames[\'top\'].HoverSel(0,\'_\',\'_\');window.status=\'Dynamic Menu\';\" OnClick=\"frames[\'top\'].execURL(\'\', \'frames[top]\');\">Dynamic Menu</td></tr></table></div>");
			close();
		}
	if(NS)
		with(document) {
			open();
			write("<layer name=\"menu1\" background=\"\" top=\"0%\" left=\"0%\" clip=\"0,0,109,38\" z-index=100  bgColor=\"#C0C0C0\" visibility=\"hidden\"><layer name=MC1EH1 top=2 left=0 width=109 height=16 z-index=101 OnMouseOver=\"frames[\'top\'].NSHoverSel(0,mFrame.document.layers[\'menu1\'].layers[\'MC1\'],\'#0000FF\',105,16);window.status=\'Menu 1 Option 1\';\"></layer><layer name=MC1 top=2 left=2 width=105 height=16 z-index=100  bgcolor=\"#C0C0C0\"><p align=left><font face=Tahoma point-size=9 color=#000000>Menu 1 Option 1</font></layer><layer name=MC2EH2 top=19 left=0 width=109 height=16 z-index=101 OnMouseOver=\"frames[\'top\'].NSHoverSel(0,mFrame.document.layers[\'menu1\'].layers[\'MC2\'],\'#0000FF\',105,16);window.status=\'\';\"></layer><layer name=MC2 top=19 left=2 width=105 height=16 z-index=100  bgcolor=\"#C0C0C0\"><p align=left><font face=Tahoma point-size=9 color=#000000>Menu 1 Option 2</font></layer></layer><layer name=\"menu2\" background=\"\" top=\"0%\" left=\"0%\" clip=\"0,0,95,21\" z-index=100  bgColor=\"#C0C0C0\" visibility=\"hidden\"><layer name=MC3EH3 top=2 left=0 width=95 height=16 z-index=101 OnMouseOver=\"frames[\'top\'].NSHoverSel(0,mFrame.document.layers[\'menu2\'].layers[\'MC3\'],\'#0000FF\',91,16);window.status=\'Dynamic Menu\';\"></layer><layer name=MC3 top=2 left=2 width=91 height=16 z-index=100  bgcolor=\"#C0C0C0\"><p align=left><font face=Tahoma point-size=9 color=#000000>Dynamic Menu</font></layer></layer>");
			close();
		}
SetUpEvents();



