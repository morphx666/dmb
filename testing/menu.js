////////////////////////////////////////////////
//  DHTML MENU BUILDER 2.6.303                //
//  (c)xFX JumpStart                          //
//                                            //
//  PSN: 18E3-110316-XFX-2366                 //
//                                            //
//  GENERATED: 10/5/2000 - 2:49:51 AM         //
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
	var xOff = 0;

	var HideSpeed = 200;

var BV=parseInt(navigator.appVersion);
var BN=window.navigator.appName;
var IsMac=(navigator.userAgent.indexOf('Mac')!=-1)?true:false;
var Opera=(navigator.userAgent.indexOf('Opera')!=-1)?true:false;
var NS=(BN.indexOf('Netscape')!=-1&&(BV==4)&&!Opera)?true:false;
var IE=(BN.indexOf('Explorer')!=-1&&(BV>=4)&&!Opera)?true:false;
IE=true;NS=false;

	if ((frames.length==0) && IsMac)
		frames.top = window;

	if(IE)
		xOff = 2;
	cFrame = eval(frames['self']);

	var fx = Math.round(Math.random()*3);



	nStyle[0]="border: 1px;  font-family: Tahoma; font-size: 9pt; font-weight: bold; font-style: none; color: #DFDFDF; background-color: #808080; cursor: default";
	hStyle[0]="border: 1px;  font-family: Tahoma; font-size: 9pt; font-weight: bold; font-style: none; color: #FFFFFF; background-color: #0080C0; cursor: default;";
	nStyle[1]="border: 1px;  font-family: Tahoma; font-size: 9pt; font-weight: bold; font-style: none; color: #DFDFDF; background-color: #808080; cursor: default";
	hStyle[1]="border: 1px;  font-family: Tahoma; font-size: 9pt; font-weight: bold; font-style: none; color: #FFFFFF; background-color: #0080C0; cursor: default;";
	nStyle[2]="border: 1px;  font-family: Tahoma; font-size: 9pt; font-weight: bold; font-style: none; color: #DFDFDF; background-color: #808080; cursor: default";
	hStyle[2]="border: 1px;  font-family: Tahoma; font-size: 9pt; font-weight: bold; font-style: none; color: #FFFFFF; background-color: #0080C0; cursor: default;";
	nStyle[3]="border: 1px;  font-family: Tahoma; font-size: 9pt; font-weight: bold; font-style: none; color: #DFDFDF; background-color: #808080; cursor: default";
	hStyle[3]="border: 1px;  font-family: Tahoma; font-size: 9pt; font-weight: bold; font-style: none; color: #FFFFFF; background-color: #0080C0; cursor: default;";
	nStyle[4]="border: 1px;  font-family: Tahoma; font-size: 9pt; font-weight: bold; font-style: none; color: #DFDFDF; background-color: #808080; cursor: default";
	hStyle[4]="border: 1px;  font-family: Tahoma; font-size: 9pt; font-weight: bold; font-style: none; color: #FFFFFF; background-color: #0080C0; cursor: default;";
	nStyle[5]="border: 1px;  font-family: Tahoma; font-size: 9pt; font-weight: bold; font-style: none; color: #DFDFDF; background-color: #808080; cursor: default";
	hStyle[5]="border: 1px;  font-family: Tahoma; font-size: 9pt; font-weight: bold; font-style: none; color: #FFFFFF; background-color: #0080C0; cursor: default;";
	nStyle[6]="border: 1px;  font-family: Tahoma; font-size: 9pt; font-weight: bold; font-style: none; color: #DFDFDF; background-color: #808080; cursor: default";
	hStyle[6]="border: 1px;  font-family: Tahoma; font-size: 9pt; font-weight: bold; font-style: none; color: #FFFFFF; background-color: #0080C0; cursor: default;";
	nStyle[7]="border: 1px;  font-family: Tahoma; font-size: 9pt; font-weight: bold; font-style: none; color: #DFDFDF; background-color: #808080; cursor: default";
	hStyle[7]="border: 1px;  font-family: Tahoma; font-size: 9pt; font-weight: bold; font-style: none; color: #FFFFFF; background-color: #0080C0; cursor: default;";
	nStyle[8]="border: 1px;  font-family: Tahoma; font-size: 9pt; font-weight: bold; font-style: none; color: #DFDFDF; background-color: #808080; cursor: default";
	hStyle[8]="border: 1px;  font-family: Tahoma; font-size: 9pt; font-weight: bold; font-style: none; color: #FFFFFF; background-color: #0080C0; cursor: default;";
	nTCode[1]="frames[\'self\'].execURL(\'javascript:alert(These options are for demonstration purposes only);\', \'frames[self]\');";
	nLayer[1]="<p align=left><b><font face=Tahoma point-size=9 color=#DFDFDF>Home</font></b>";
	hLayer[1]="<p align=left><b><font face=Tahoma point-size=9 color=#FFFFFF>Home</font></b>";
	nTCode[2]="frames[\'self\'].ShowMenu(\'ProdCats\', 0, 0, true)";
	nLayer[2]="<p align=left><b><font face=Tahoma point-size=9 color=#DFDFDF><img align=\"right\" name=\"mnuNavProductsRImg\" src=\"arrow.gif\" width=\"16\" height=\"16\" border=\"0\">Products</font></b>";
	hLayer[2]="<p align=left><b><font face=Tahoma point-size=9 color=#FFFFFF><img align=\"right\" name=\"mnuNavProductsRImg\" src=\"arrow.gif\" width=\"16\" height=\"16\" border=\"0\">Products</font></b>";
	nTCode[3]="frames[\'self\'].execURL(\'javascript:alert(These options are for demonstration purposes only);\', \'frames[self]\');";
	nLayer[3]="<p align=left><b><font face=Tahoma point-size=9 color=#DFDFDF>Download</font></b>";
	hLayer[3]="<p align=left><b><font face=Tahoma point-size=9 color=#FFFFFF>Download</font></b>";
	nTCode[4]="frames[\'self\'].execURL(\'javascript:alert(These options are for demonstration purposes only);\', \'frames[self]\');";
	nLayer[4]="<p align=left><b><font face=Tahoma point-size=9 color=#DFDFDF>Buy Online</font></b>";
	hLayer[4]="<p align=left><b><font face=Tahoma point-size=9 color=#FFFFFF>Buy Online</font></b>";
	nTCode[5]="frames[\'self\'].execURL(\'javascript:alert(These options are for demonstration purposes only);\', \'frames[self]\');";
	nLayer[5]="<p align=left><b><font face=Tahoma point-size=9 color=#DFDFDF>Support</font></b>";
	hLayer[5]="<p align=left><b><font face=Tahoma point-size=9 color=#FFFFFF>Support</font></b>";
	nTCode[6]="frames[\'self\'].execURL(\'javascript:alert(These options are for demonstration purposes only);\', \'frames[self]\');";
	nLayer[6]="<p align=left><b><font face=Tahoma point-size=9 color=#DFDFDF>Older News</font></b>";
	hLayer[6]="<p align=left><b><font face=Tahoma point-size=9 color=#FFFFFF>Older News</font></b>";
	nTCode[7]="frames[\'self\'].execURL(\'javascript:alert(These options are for demonstration purposes only);\', \'frames[self]\');";
	nLayer[7]="<p align=left><b><font face=Tahoma point-size=9 color=#DFDFDF>Utilities</font></b>";
	hLayer[7]="<p align=left><b><font face=Tahoma point-size=9 color=#FFFFFF>Utilities</font></b>";
	nTCode[8]="frames[\'self\'].execURL(\'javascript:alert(These options are for demonstration purposes only);\', \'frames[self]\');";
	nLayer[8]="<p align=left><b><font face=Tahoma point-size=9 color=#DFDFDF>Audio</font></b>";
	hLayer[8]="<p align=left><b><font face=Tahoma point-size=9 color=#FFFFFF>Audio</font></b>";
	nTCode[9]="frames[\'self\'].execURL(\'javascript:alert(These options are for demonstration purposes only);\', \'frames[self]\');";
	nLayer[9]="<p align=left><b><font face=Tahoma point-size=9 color=#DFDFDF>ActiveX Controls</font></b>";
	hLayer[9]="<p align=left><b><font face=Tahoma point-size=9 color=#FFFFFF>ActiveX Controls</font></b>";
	var mnuNavProductsRImgOff = new Image;
	var mnuNavProductsRImgOn = new Image;
	mnuNavProductsRImgOff.src = 'arrow.gif';
	mnuNavProductsRImgOn.src = 'arrow.gif';


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
//Version 9.3
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
					return false;
			if(OpenMenus[nOM].SelCommand || nOM>1)
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

		return true;
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
//Version 10.5
//
		var n;
		var LayerHTM;
		if(mode==0 && OpenMenus[nOM].SelCommand!=null)
			NSHoverSel(1);
		
		if(mode==0) {
			n = parseInt(mc.name.substr(2));
			if(nOM>1)
				if(mc==OpenMenus[nOM-1].SelCommand)
					return false;
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

		return true;
	}

	function Hide() {
//IE,NS
//This function hides the last opened group and it keeps hiding all the groups until
//no more groups are opened or the mouse is over one of them.
//Also takes care of reseting any highlighted commands.
//------------------------------
//Version 3.2
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

		if(nOM==0)
			FormsTweak("visible");
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
//Version 14.0
//
		alert(1);
		x = parseInt(x);
		y = parseInt(y);
		if(AnimHnd && nOM>0) {
			AnimStep=100;
			Animate();
		}
		if(IE)
			var Menu = mFrame.document.all[mName];
		if(NS)
			var Menu = mFrame.document.layers[mName];
		if(!Menu)
			return false;
		if(IE)
			Menu = Menu.style;
		if(Menu==OpenMenus[nOM] || HTHnd[nOM])
			return false;
		
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
				Menu.left = (x+parseInt(Menu.width)>pW)?parseInt(OpenMenus[nOM].left) - parseInt(Menu.width) + 6:x;
				Menu.top =  (y+parseInt(Menu.height)>pH)?pH - parseInt(Menu.height):y;
			} else {
				Menu.left = (x+parseInt(Menu.width)>pW)?pW - parseInt(Menu.width):x;
				Menu.top =  (y+parseInt(Menu.height)>pH)?pH - parseInt(Menu.height):y;
			}
			if(!IsMac)
				Menu.clip = "rect(0 0 0 0)";
		}
		if(NS) {
			if(isCascading) {
				x = OpenMenus[nOM].left + OpenMenus[nOM].clip.width - 6;
				y = OpenMenus[nOM].top + OpenMenus[nOM].SelCommand.top;
				x = (x+Menu.w>pW)?OpenMenus[nOM].left - Menu.w + 6:x;
				y = (y+Menu.h>pH)?pH - Menu.h:y;
			} else {
				x = (x+Menu.w>pW)?pW - Menu.w:x;
				y = (y+Menu.h>pH)?pH - Menu.h:y;
			}
			Menu.clip.width = 0;
			Menu.clip.height = 0;
			Menu.moveToAbsolute(x,y);
		}
		if(isCascading)
			Menu.zIndex = parseInt(OpenMenus[nOM].zIndex) + 1;
		Menu.visibility = "visible";
		OpenMenus[++nOM] = Menu;
		HTHnd[nOM] = 0;
		if((IE && !IsMac) || NS)
			AnimHnd = window.setTimeout("Animate()", 10);
		FormsTweak("hidden");

		return true;
	}

	function Animate() {
//IE,NS
//This function is called by ShowMenu every time a new group must be displayed and produces the predefined unfolding effect.
//Currently is disabled for Navigator, because of some weird bugs we found with the clip property of the layers.
//------------------------------
//Version 1.9
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
//Version 1.6
//
		var m = OpenMenus[nOM];
		if(!m)
			return false;
		if(IE&&BV==4)
			SetPointerPos();
		var l = parseInt(m.left) + xOff;
		var r = l+((IE)?parseInt(m.width):m.clip.width) - xOff;
		var t = parseInt(m.top) + xOff;
		var b = t+((IE)?parseInt(m.height):m.clip.height) - xOff;
		return ((mX>=l && mX<=r) && (mY>=t && mY<=b));
	}

	function SetPointerPos(e) {
//IE,NS
//This function sets the mX and mY variables with the current position of the mouse pointer.
//------------------------------
//e: Only used under Navigator, corresponds to the Event object.
//------------------------------
//Version 1.0
//
		if(IE) {
			if(event==null)
				if(mFrame.window.event==null)
					return;
				else
					e = mFrame.window.event;
			else
				e = event;
			mX = e.clientX + mFrame.document.body.scrollLeft;
			mY = e.clientY + mFrame.document.body.scrollTop;
		}
		if(NS) {
			mX = e.pageX;
			mY = e.pageY;
		}
	}
	

	function HideMenus(e) {
//IE,NS
//This function checks if the mouse pointer is on a valid position and if the current menu should be kept visible.
//The function is called every time the mouse pointer is moved over the document area.
//------------------------------
//e: Only used under Navigator, corresponds to the Event object.
//------------------------------
//Version 24.3
//
		SetPointerPos(e);
		if(nOM>0)
			if(OpenMenus[nOM].SelCommand!=null)
				while(!InMenu() && !HTHnd[nOM]) {
					HTHnd[nOM] = window.setTimeout("Hide()", HideSpeed);
					if(nOM==0)
						break;
				}
	}
	
	function FormsTweak(state) {
//IE
//This is an undocumented function, which can be used to hide every listbox (or combo) element on a page.
//This can be useful if the menus will be displayed over an area where is a combo box, which is an element that cannot be placed behind the menus and it will always appear over the menus resulting in a very undesirable effect.
//------------------------------
//Version 2.0
//
		if(DoFormsTweak && IE)
			for(var f = 0; f <= (mFrame.document.forms.length - 1); f++)
				for(var e = 0; e <= (mFrame.document.forms[f].elements.length - 1); e++)
					if(mFrame.document.forms[f].elements[e].type=="select-one")
						mFrame.document.forms[f].elements[e].style.visibility = state;
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
//Version 1.2
//
		if(typeof(mFrame)=="undefined")
		mFrame = eval(frames['self']);
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
//Version 2.1
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
			write("<div id=\"mnuNav\" style=\"position: absolute; top: 0%; left: 0%; width: 92; height: 150; z-index: 100; visibility: hidden;\"><table id=\"dmbMenu\" background=\"\" border=\"0\" cellpadding=\"2\" style=\"background-color: #C0C0C0; color: #000000; border-left: #C0C0C0 solid 1; border-right: 1 solid #C0C0C0; border-top: 1 solid #C0C0C0; border-bottom: 1 solid #C0C0C0\" width=\"92\"><tr><td nowrap height=20 align=\"left\" style=\"border: 1px; font-family: Tahoma; font-size: 9pt; font-weight: bold; font-style: none; color: #DFDFDF; cursor: default; background-color: #808080;\" id=0 OnMouseOver=\"frames[\'self\'].HoverSel(0,\'_\',\'_\');window.status=\'Return to the Home or General News page\';\" OnClick=\"frames[\'self\'].execURL(\'javascript:alert(These options are for demonstration purposes only);\', \'frames[self]\');\">Home</td></tr><tr><td nowrap height=22 align=\"left\" style=\"border: 1px; font-family: Tahoma; font-size: 9pt; font-weight: bold; font-style: none; color: #DFDFDF; cursor: default; background-color: #808080;\" id=1 OnMouseOver=\"frames[\'self\'].HoverSel(0,\'_\',\'mnuNavProductsRImg\');window.status=\'Products\';\" OnClick=\"frames[\'self\'].ShowMenu(\'ProdCats\', 0, 30, true);\"><img align=\"right\" name=\"mnuNavProductsRImg\" src=\"arrow.gif\" width=\"16\" height=\"16\" border=\"0\">Products</td></tr><tr><td nowrap height=20 align=\"left\" style=\"border: 1px; font-family: Tahoma; font-size: 9pt; font-weight: bold; font-style: none; color: #DFDFDF; cursor: default; background-color: #808080;\" id=2 OnMouseOver=\"frames[\'self\'].HoverSel(0,\'_\',\'_\');window.status=\'Download\';\" OnClick=\"frames[\'self\'].execURL(\'javascript:alert(These options are for demonstration purposes only);\', \'frames[self]\');\">Download</td></tr><tr><td nowrap height=20 align=\"left\" style=\"border: 1px; font-family: Tahoma; font-size: 9pt; font-weight: bold; font-style: none; color: #DFDFDF; cursor: default; background-color: #808080;\" id=3 OnMouseOver=\"frames[\'self\'].HoverSel(0,\'_\',\'_\');window.status=\'Buy our programs on line using a secure transaction system\';\" OnClick=\"frames[\'self\'].execURL(\'javascript:alert(These options are for demonstration purposes only);\', \'frames[self]\');\">Buy Online</td></tr><tr><td nowrap height=20 align=\"left\" style=\"border: 1px; font-family: Tahoma; font-size: 9pt; font-weight: bold; font-style: none; color: #DFDFDF; cursor: default; background-color: #808080;\" id=4 OnMouseOver=\"frames[\'self\'].HoverSel(0,\'_\',\'_\');window.status=\'Support Form\';\" OnClick=\"frames[\'self\'].execURL(\'javascript:alert(These options are for demonstration purposes only);\', \'frames[self]\');\">Support</td></tr><tr><td nowrap height=20 align=\"left\" style=\"border: 1px; font-family: Tahoma; font-size: 9pt; font-weight: bold; font-style: none; color: #DFDFDF; cursor: default; background-color: #808080;\" id=5 OnMouseOver=\"frames[\'self\'].HoverSel(0,\'_\',\'_\');window.status=\'View news more than a month old, in case you have missed them\';\" OnClick=\"frames[\'self\'].execURL(\'javascript:alert(These options are for demonstration purposes only);\', \'frames[self]\');\">Older News</td></tr></table></div><div id=\"ProdCats\" style=\"position: absolute; top: 0%; left: 0%; width: 122; height: 76; z-index: 100; visibility: hidden;\"><table id=\"dmbMenu\" background=\"\" border=\"0\" cellpadding=\"2\" style=\"background-color: #C0C0C0; color: #000000; border-left: #C0C0C0 solid 1; border-right: 1 solid #C0C0C0; border-top: 1 solid #C0C0C0; border-bottom: 1 solid #C0C0C0\" width=\"122\"><tr><td nowrap height=20 align=\"left\" style=\"border: 1px; font-family: Tahoma; font-size: 9pt; font-weight: bold; font-style: none; color: #DFDFDF; cursor: default; background-color: #808080;\" id=6 OnMouseOver=\"frames[\'self\'].HoverSel(0,\'_\',\'_\');window.status=\'Utilities\';\" OnClick=\"frames[\'self\'].execURL(\'javascript:alert(These options are for demonstration purposes only);\', \'frames[self]\');\">Utilities</td></tr><tr><td nowrap height=20 align=\"left\" style=\"border: 1px; font-family: Tahoma; font-size: 9pt; font-weight: bold; font-style: none; color: #DFDFDF; cursor: default; background-color: #808080;\" id=7 OnMouseOver=\"frames[\'self\'].HoverSel(0,\'_\',\'_\');window.status=\'Audio\';\" OnClick=\"frames[\'self\'].execURL(\'javascript:alert(These options are for demonstration purposes only);\', \'frames[self]\');\">Audio</td></tr><tr><td nowrap height=20 align=\"left\" style=\"border: 1px; font-family: Tahoma; font-size: 9pt; font-weight: bold; font-style: none; color: #DFDFDF; cursor: default; background-color: #808080;\" id=8 OnMouseOver=\"frames[\'self\'].HoverSel(0,\'_\',\'_\');window.status=\'ActiveX Controls\';\" OnClick=\"frames[\'self\'].execURL(\'javascript:alert(These options are for demonstration purposes only);\', \'frames[self]\');\">ActiveX Controls</td></tr></table></div>");
			close();
		}
	if(NS)
		with(document) {
			open();
			write("<layer name=\"mnuNav\" background=\"\" top=0 left=0 clip=\"0,0,90,126\" z-index=100  bgColor=\"#C0C0C0\" visibility=\"hidden\"><layer name=MC1EH1 top=2 left=0 width=90 height=18 z-index=101 OnMouseOver=\"frames[\'self\'].NSHoverSel(0,mFrame.document.layers[\'mnuNav\'].layers[\'MC1\'],\'#0080C0\',86,18);window.status=\'Return to the Home or General News page\';\"></layer><layer name=MC1 top=2 left=2 width=86 height=18 z-index=100  bgcolor=\"#808080\"><p align=left><b><font face=Tahoma point-size=9 color=#DFDFDF>Home</font></b></layer><layer name=MC2EH2 top=22 left=0 width=90 height=20 z-index=101 OnMouseOver=\"frames[\'self\'].NSHoverSel(0,mFrame.document.layers[\'mnuNav\'].layers[\'MC2\'],\'#0080C0\',86,20);window.status=\'Products\';\"></layer><layer name=MC2 top=22 left=2 width=86 height=20 z-index=100  bgcolor=\"#808080\"><p align=left><b><font face=Tahoma point-size=9 color=#DFDFDF><img align=\"right\" name=\"mnuNavProductsRImg\" src=\"arrow.gif\" width=\"16\" height=\"16\" border=\"0\">Products</font></b></layer><layer name=MC3EH3 top=44 left=0 width=90 height=18 z-index=101 OnMouseOver=\"frames[\'self\'].NSHoverSel(0,mFrame.document.layers[\'mnuNav\'].layers[\'MC3\'],\'#0080C0\',86,18);window.status=\'Download\';\"></layer><layer name=MC3 top=44 left=2 width=86 height=18 z-index=100  bgcolor=\"#808080\"><p align=left><b><font face=Tahoma point-size=9 color=#DFDFDF>Download</font></b></layer><layer name=MC4EH4 top=64 left=0 width=90 height=18 z-index=101 OnMouseOver=\"frames[\'self\'].NSHoverSel(0,mFrame.document.layers[\'mnuNav\'].layers[\'MC4\'],\'#0080C0\',86,18);window.status=\'Buy our programs on line using a secure transaction system\';\"></layer><layer name=MC4 top=64 left=2 width=86 height=18 z-index=100  bgcolor=\"#808080\"><p align=left><b><font face=Tahoma point-size=9 color=#DFDFDF>Buy Online</font></b></layer><layer name=MC5EH5 top=84 left=0 width=90 height=18 z-index=101 OnMouseOver=\"frames[\'self\'].NSHoverSel(0,mFrame.document.layers[\'mnuNav\'].layers[\'MC5\'],\'#0080C0\',86,18);window.status=\'Support Form\';\"></layer><layer name=MC5 top=84 left=2 width=86 height=18 z-index=100  bgcolor=\"#808080\"><p align=left><b><font face=Tahoma point-size=9 color=#DFDFDF>Support</font></b></layer><layer name=MC6EH6 top=104 left=0 width=90 height=18 z-index=101 OnMouseOver=\"frames[\'self\'].NSHoverSel(0,mFrame.document.layers[\'mnuNav\'].layers[\'MC6\'],\'#0080C0\',86,18);window.status=\'View news more than a month old, in case you have missed them\';\"></layer><layer name=MC6 top=104 left=2 width=86 height=18 z-index=100  bgcolor=\"#808080\"><p align=left><b><font face=Tahoma point-size=9 color=#DFDFDF>Older News</font></b></layer></layer><layer name=\"ProdCats\" background=\"\" top=0 left=0 clip=\"0,0,120,64\" z-index=100  bgColor=\"#C0C0C0\" visibility=\"hidden\"><layer name=MC7EH7 top=2 left=0 width=120 height=18 z-index=101 OnMouseOver=\"frames[\'self\'].NSHoverSel(0,mFrame.document.layers[\'ProdCats\'].layers[\'MC7\'],\'#0080C0\',116,18);window.status=\'Utilities\';\"></layer><layer name=MC7 top=2 left=2 width=116 height=18 z-index=100  bgcolor=\"#808080\"><p align=left><b><font face=Tahoma point-size=9 color=#DFDFDF>Utilities</font></b></layer><layer name=MC8EH8 top=22 left=0 width=120 height=18 z-index=101 OnMouseOver=\"frames[\'self\'].NSHoverSel(0,mFrame.document.layers[\'ProdCats\'].layers[\'MC8\'],\'#0080C0\',116,18);window.status=\'Audio\';\"></layer><layer name=MC8 top=22 left=2 width=116 height=18 z-index=100  bgcolor=\"#808080\"><p align=left><b><font face=Tahoma point-size=9 color=#DFDFDF>Audio</font></b></layer><layer name=MC9EH9 top=42 left=0 width=120 height=18 z-index=101 OnMouseOver=\"frames[\'self\'].NSHoverSel(0,mFrame.document.layers[\'ProdCats\'].layers[\'MC9\'],\'#0080C0\',116,18);window.status=\'ActiveX Controls\';\"></layer><layer name=MC9 top=42 left=2 width=116 height=18 z-index=100  bgcolor=\"#808080\"><p align=left><b><font face=Tahoma point-size=9 color=#DFDFDF>ActiveX Controls</font></b></layer></layer>");
			close();
		}
SetUpEvents();



