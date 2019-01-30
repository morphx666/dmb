var BV=parseInt(navigator.appVersion.indexOf("MSIE")>0?navigator.appVersion.split(";")[1].substr(6):navigator.appVersion);
var BN=window.navigator.appName;
var IsWin=(navigator.userAgent.indexOf('Windows')!=-1)?true:false;
var OP=(navigator.userAgent.indexOf('Opera')!=-1&&BV>=4)?true:false;
var NS=(BN.indexOf('Netscape')!=-1&&(BV==4)&&!OP)?true:false;
var SM=(BN.indexOf('Netscape')!=-1&&(BV>=5)||OP)?true:false;
var IE=(BN.indexOf('Explorer')!=-1&&(BV>=4)||SM)?true:false;

var mX;
var mY;

function GetDIV(dname) {
	if(SM||OP)
		return document.getElementById(dname);
	else
		return document.all[dname];
}

function GetLeftTop() {
	return [OP?0:SM?scrollX:document.body.scrollLeft,OP?0:SM?scrollY:document.body.scrollTop];
}

function GetWidthHeight() {
	if(IE&&!SM)
		return [document.body.clientWidth,document.body.clientHeight];
	if(!IE||SM)
		return [innerWidth,innerHeight];
}

function SetPointerPos(e) {
	if(IE&&!SM) {
		e = window.event;
	}
	mX = e.clientX + GetLeftTop()[0];
	mY = e.clientY + GetLeftTop()[1];
}

function HideIt() {
	if(p)
		p.style.visibility = "hidden";
}

function say(what) {
	var info = GetDIV("Info"+what);
	var pW = GetWidthHeight()[0];
	var l;
	
	if(!info) info = GetDIV("NullInfo");

	p.innerHTML = info.innerHTML	
	with(p.style) {
		width = "auto";
		height = "auto";
		overflow = "auto";
	}
	with(p.style) {
		l = mX - 60;
		top = mY + 10 + "px";
		cursor = "default";
		visibility = "visible";
		width = parseInt(p.offsetWidth) + 10 + "px";
		if(parseInt(width)>(pW-18)) width = pW - 18 + "px";
		height = parseInt(p.offsetHeight);
		
		if((l + parseInt(width)) > (pW-18)) l = pW - 18 - parseInt(width);
		if(l<0) l = 0;
		left = l + "px";
		
		if(parseInt(height)>300) {
			height = "300px";
			overflow = "scroll";
		}
	}
}

var p = GetDIV("popup");

document.onmousemove = SetPointerPos;
document.onclick = HideIt;