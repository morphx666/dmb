var nTCode=new Array;var AnimStep=0;var AnimHnd=0;var NSDelay=0;var MenusReady=false;var smHnd=0;var mhdHnd=0;var lsc=null;var imgObj=null;var IsContext=false;var IsFrames=false;var dxFilter=null;var AnimSpeed=35;var TimerHideDelay=104;var smDelay=71;var rmDelay=15;var scDelay=0;var cntxMenu='';var DoFormsTweak=true;var mibc=0;var mibm;var mibs=50;var nsOWH;var mFrame;var cFrame=self;var om=new Array;var nOM=0;var mX;var mY;var BV=parseFloat(navigator.appVersion.indexOf("MSIE")>0?navigator.appVersion.split(";")[1].substr(6):navigator.appVersion);var BN=navigator.appName;var nua=navigator.userAgent;var IsWin=(nua.indexOf('Win')!=-1);var IsMac=(nua.indexOf('Mac')!=-1);var KQ=(BN.indexOf('Konqueror')!=-1&&(BV>=5))||(nua.indexOf('Safari')!=-1);var OP=(nua.indexOf('Opera')!=-1&&BV>=4);var NS=(BN.indexOf('Netscape')!=-1&&(BV>=4&&BV<5)&&!OP);var SM=(BN.indexOf('Netscape')!=-1&&(BV>=5)||OP);var IE=(BN.indexOf('Explorer')!=-1&&(BV>=4)||SM||KQ);var IX=(IE&&IsWin&&!SM&&!OP&&(BV>=5.5)&&(dxFilter!=null)&&(nua.indexOf('CE')==-1));if(!eval(frames['self'])){frames.self=window;frames.top=top;}var lmcHS=null;var tbNum=0;var fx=0;nTCode[12]="cFrame.execURL('http://www.google.com','_self');";nTCode[13]="cFrame.f31();cFrame.dmbNW=window.open(cFrame._purl('http://software.xfx.net'),'NewWindow','left=80,top=80,width=600,height=400,directories=0,channelmode=0,toolbar=0,fullscreen=0,location=0,menubar=0,resizable=1,scrollbars=1,status=0,titlebar=1');if(cFrame.IsFrames&&IE&&!SM)mFrame.location.reload();cFrame.dmbNW.focus();";function f6(id){var mc=f34(id);if(nOM>1){if(mc==om[nOM-1].sc) return false;while(true){if(!nOM) return false;if(gpid(mc)==om[nOM].id) break;Hide();}if(nOM&&scDelay) mc.onmouseover();}}function f13(mode,mc){var mcN;var mn;f16("smHnd");if(!nOM) return false;if(mode==0&&om[nOM].sc!=null) f13(1);if(mode==0){mn=mc.name.substr(0,mc.name.indexOf("EH"));mcN=mc.parentLayer.layers[mn+"N"];mcN.mcO=mc.parentLayer.layers[mn+"O"];if(nOM>1) if(mc==om[nOM-1].sc) return false;while(!f32()&&nOM>1) Hide();om[nOM].sc=mcN;mcN.mcO.visibility="show";mcN.visibility="hide";}else{mcN=(mode==1)?om[nOM].sc:om[nOM].op;mcN.visibility="show";mcN.mcO.visibility="hide";om[nOM].sc=null;}return true;}function Hide(chk){var m;var cl=false;f16("mhdHnd");f16("AnimHnd");if(chk)if(f32()) return false;if(nOM){m=om[nOM];if(m.sc!=null){if(IE) HoverSel(1);if(NS) f13(1);}if(m.op!=null){if(IE) HoverSel(3);if(NS) f13(3);}f15(m,"hidden");f16("om[nOM].hsm");nOM--;cl=(nOM==0);if(cl) imgObj=null;}if(cl||chk){f16("smHnd");if(tbNum&&lmcHS) if(!lmcHS.disable){if(IE) hsHoverSel(1);if(NS) hsNSHoverSel(1);}if(!lmcHS) window.status="";}if((nOM>0||lmcHS)&&!f22()) mhdHnd=window.setTimeout("Hide(1)",TimerHideDelay/20);return true;}function f15(m,s){if(IX) if(document.readyState=="complete"&&m.getAttribute("filters")!=null){if(!m.fs){m.fsn=m.filters.length;m.style.filter=dxFilter+m.style.filter;m.fs=true;}for(var i=0;i<m.filters.length-m.fsn;i++){m.filters[i].apply();m.style.visibility=s;m.filters[i].play();}}m.style.visibility=s;if(s=="hidden"&&!IX&&!NS) m.style.left=m.style.top="0px";f17(s=="visible"?"hidden":"visible");}function ShowMenu(mName,x,y,isc,hsimg,algn){if(!algn) algn=0;var f=["f19('"+mName+"',"+x+","+y+","+isc+",'"+hsimg+"',"+algn+")"];f16("smHnd");if(isc){if(nOM==0) return false;lsc=om[nOM].sc;f[1]="if(nOM)if(lsc==om[nOM].sc)";f[2]=smDelay;}else{if(nOM>0) if(om[1].id==mName) return false;f16("mhdHnd");if(algn<0&&!lmcHS) return false;f[1]="f31(1);";f[2]=f33();}smHnd=window.setTimeout(f[1]+f[0],f[2]);return true;}function f19(mName,x,y,isc,hsimg,algn){var xy;x=pri(x);y=pri(y);var Menu=f34(mName);if(!Menu) return false;if(Menu==om[nOM]) return false;if(NS) Menu.style=Menu;if(AnimHnd&&nOM>0){AnimStep=100;f30();}Menu.op=nOM>0?om[nOM].sc:null;Menu.sc=null;imgObj=null;if(isc){if(!NS) f6(om[nOM].sc.id);xy=f2(Menu,algn);var gs=om[nOM].gs;if(gs) xy[1]+=pri(gs.top);}else{Menu.pHS=lmcHS;xy=(algn<0?GetHSPos(x,y,NS?Menu.w:Menu.offsetWidth,NS?Menu.h:Menu.offsetHeight,-algn):[x,y]);if(hsimg){var hss=hsimg.split("|");imgObj=NS?f18(cFrame.document,hss[0]):cFrame.document.images[hss[0]];if(imgObj){if(hss[1]&2) xy[0]=f29(Menu.style,imgObj,algn)[0]+(IsFrames?f12()[0]:0)+f23()[0];if(hss[1]&1) xy[1]=f29(Menu.style,imgObj,algn)[1]+(IsFrames?f12()[1]:0)+f23()[1];}}}if(xy){x=xy[0];y=xy[1];}var pWH=[f1()[0]+f12()[0],f1()[1]+f12()[1]];if(IE) with(Menu.style){if(SM) display="none";left=f35(x,pri(width),pWH[0],0)+"px";top=f35(y,pri(height),pWH[1],1)+"px";if(!IX&&!SM&&IsWin) clip="rect(0 0 0 0)";}if(NS){Menu.clip.width=0;Menu.clip.height=0;Menu.moveToAbsolute(f35(x,Menu.w,pWH[0]),f35(y,Menu.h,pWH[1]));}Menu.style.zIndex=1000+tbNum+nOM;om[++nOM]=Menu;if(!NS) FixCommands(mName);if(SM) Menu.style.display="inline";if(!IX){if((IE&&IsWin&&!SM)||(NS&&Menu.style.clip.width==0)) AnimHnd=window.setTimeout("f30()",10);}f15(Menu,"visible");IsContext=false;smHnd=0;return true;}function f33(){return rmDelay/(nOM==0&&scDelay>0?4:1);}function f23(f){var mo=[0,0];if(!f) f=mFrame;if(IsMac&&IE&&!SM&&!KQ&&(BV>=5)) mo=[pri(f.document.body.leftMargin),pri(f.document.body.topMargin)];return mo;}function f2(mg,a){var x;var y;var pg=om[nOM];var sc=(NS?pg.sc:f34("O"+pg.sc.id.substr(1)));if(NS){pg.offsetLeft=pg.left;pg.offsetTop=pg.top;pg.offsetWidth=pg.w;pg.offsetHeight=pg.h;mg.offsetWidth=mg.w;mg.offsetHeight=mg.h;sc.offsetLeft=sc.left;sc.offsetTop=sc.top;sc.offsetWidth=sc.clip.width;sc.offsetHeight=sc.clip.height;}var lp=pg.offsetLeft+sc.offsetLeft;var tp=pg.offsetTop+sc.offsetTop;switch(a){case 6:x=lp+sc.offsetWidth;y=tp;break;}return [x,y];}function f30(){var r='';var nw=nh=0;if(AnimStep+AnimSpeed>100) AnimStep=100;switch(fx){case 1:if(IE) r="0 "+AnimStep+"% "+AnimStep+"% 0";if(NS) nw=AnimStep;nh=AnimStep;break;case 2:if(IE) r="0 100% "+AnimStep+"% 0";if(NS) nw=100;nh=AnimStep;break;case 3:if(IE) r="0 "+AnimStep+"% 100% 0";if(NS) nw=AnimStep;nh=100;break;case 0:if(IE) r="0 100% 100% 0";if(NS) nw=100;nh=100;break;}if(om[nOM]){with(om[nOM].style){if(IE) clip="rect("+r+")";if(NS){clip.width=w*(nw/100);clip.height=h*(nh/100);}}AnimStep+=AnimSpeed;if(AnimStep<=100) AnimHnd=window.setTimeout("f30()",25);else{f16("AnimHnd");AnimStep=0;AnimHnd=0;}}}function f10(){var m=(imgObj?imgObj:lmcHS);var tp=[0,0];var tbb;if(!m) return false;if(imgObj) if(imgObj.name.indexOf("dmbHSdyna")!=-1){imgObj=null;return false;}var x=mX;var y=mY;if(!imgObj){if(IE){if(BV<5&&!IsMac) m.parentNode=m.parentElement;}else m.parentNode=m.parentLayer;tp=getTBPos(m.parentNode.id.substr(5));if(IE) m=m.style;}else{m.left=f28(imgObj)[0];m.top=f28(imgObj)[1];if(NS) m.clip=m;}var l=pri(m.left)+tp[0];var r=l+(IE?pri(m.width):m.clip.width);var t=pri(m.top)+tp[1];var b=t+(IE?pri(m.height):m.clip.height);if(IsFrames&&!NS){x-=f12()[0];y-=f12()[1];}return ((x>=l&&x<=r)&&(y>=t&&y<=b));}function f32(){var m=om[nOM];if(!m) return false;else if(IE) m=m.style;var l=pri(m.left);var r=l+(IE?pri(m.width):m.clip.width);var t=pri(m.top);var b=t+(IE?pri(m.height):m.clip.height);return ((mX>=l&&mX<=r)&&(mY>=t&&mY<=b));}function f4(e){if(IE){if(!SM){if(mFrame!=cFrame||event==null){var cfe=cFrame.window.event;var mfe=mFrame.window.event;if(IE&&IsMac) cfe=(cfe.type=="mousemove"?cfe:null);if(IE&&IsMac) mfe=(mfe.type=="mousemove"?mfe:null);if(mfe==null&&cfe==null) return;e=(cfe?cfe:mfe);}else e=event;}mX=e.clientX+lt[0];mY=e.clientY+lt[1];if(!KQ){f16("cFrame.iefwh");cFrame.iefwh=window.setTimeout("lt=f12()",100);}}if(NS){mX=e.pageX;mY=e.pageY;}}function f22(){var nOMb=nOM--;for(;nOM>0;nOM--) if(om[nOM].sc!=null) break;var im=f32();nOM=nOMb;return im||f32();}function f11(){return (lmcHS||imgObj?f10():(nOM==1?!(om[nOM].sc!=null):false))||(nOM>0?f22():false);}function f24(e){f4(e);if(!f11()&&mhdHnd==0) mhdHnd=window.setTimeout("mhdHnd=0;if(!f11())Hide()",TimerHideDelay);}function f17(state){var fe;if(IE&&(!SM||OP)&&DoFormsTweak){var m=om[nOM];if((BV>=5.5)&&!OP&&m&&!KQ) cIF(state=="visible"?"hidden":"visible");else if(nOM==1) for(var f=0;f<mFrame.document.forms.length;f++) for(var e=0;e<mFrame.document.forms[f].elements.length;e++){fe=mFrame.document.forms[f].elements[e];if(fe.type) if(fe.type.indexOf("select")==0) fe.style.visibility=state;}}}function execURL(url,tframe){var d=100;if(mibc&&!NS){d+=mibs * mibc;if(lmcHS) mibm=lmcHS;if(nOM) for(var n=nOM;n>0;n--) if(om[n].sc){mibm=om[n].sc;break;}if(mibm){mibm.n=mibc;f20();}}else f31();f27(escape(_purl(url)),tframe);}function f20(){var f=(mibm.id.substr(1)>1000?cFrame:mFrame);mibm.bs=!Math.abs(mibm.bs);SwapMC(mibm.bs,mibm,f);mibc--;if(mibc>=0) window.setTimeout("f20()",mibs);else{mibc=mibm.n;mibm=null;f31();}}function f27(url,tframe){var w=eval("windo"+"w.ope"+"n");url=rStr(unescape(url));if(url.indexOf("javascript:")!=url.indexOf("vbscript:")) eval(url);else{switch(tframe){case "_self":if(IE&&!SM&&(BV>4)){var a=mFrame.document.createElement("A");a.href=url;mFrame.document.body.appendChild(a);a.click();}else mFrame.location.href=url;break;case "_blank":w(url,tframe);break;default:var f=rStr(tframe);var fObj=(tframe=='_parent'?mFrame.parent:eval(f));if(typeof(fObj)=="undefined") w(url,f.substr(8,f.length-10));else fObj.location.href=url;break;}}}function rStr(s){s=xrep(s,"%1E","'");s=xrep(s,"\x1E","'");if(OP&&s.indexOf("frames[")!=-1) s=xrep(s,String.fromCharCode(s.charCodeAt(7)),"'");return xrep(s,"\x1D","\x22");}function f21(e){eval(this.TCode);}function f31(dr){var o=lmcHS;if(dr) lmcHS=null;Hide();while(nOM>0) Hide();lmcHS=o;}function tHideAll(){f16("mhdHnd");mhdHnd=window.setTimeout("mhdHnd=0;if(!f22())f31();",TimerHideDelay);}function f12(f){if(!f) f=mFrame;if(IE) if(SM) return [OP?f.pageXOffset:f.scrollX,OP?f.pageYOffset:f.scrollY];else{var b=GetBodyObj(f);return (b?[b.scrollLeft,b.scrollTop]:[0,0]);}if(NS) return [f.pageXOffset,f.pageYOffset];}function f1(f){var k=0;if(!f) f=mFrame;if(NS||SM){return [f.innerWidth,f.innerHeight];}else{var b=GetBodyObj(f);return (b?[b.clientWidth,b.clientHeight]:[0,0]);}}function f3(s){var w=[0,0];w=[(s.borderLeftStyle==""?0:pri(s.borderLeftWidth))+(s.borderRightStyle==""?0:pri(s.borderRightWidth)),(s.borderTopStyle==""?0:pri(s.borderTopWidth))+(s.borderBottomStyle==""?0:pri(s.borderBottomWidth))];return [w[0],w[1]];}function f29(m,img,arl){var x=f28(img)[0];var y=f28(img)[1];var iWH=f26(img);var mW=(NS?m.w:m.offsetWidth);var mH=(NS?m.h:m.offsetHeight);switch(arl){case 0:y+=iWH[1];break;case 1:x+=iWH[0]-mW;y+=iWH[1];break;case 2:y-=mH;break;case 3:x+=iWH[0]-mW;y-=mH;break;case 4:x-=mW;break;case 5:x-=mW;y-=mH-iWH[1];break;case 6:x+=iWH[0];break;case 7:x+=iWH[0];y-=mH-iWH[1];break;case 8:x-=mW;y+=(iWH[1]-mH)/2;break;case 9:x+=iWH[0];y+=(iWH[1]-mH)/2;break;case 10:x+=(iWH[0]-mW)/2;y-=mH;break;case 11:x+=(iWH[0]-mW)/2;y+=iWH[1];break;}return [x,y];}function f28(img){var x;var y;if(NS){y=f8(cFrame,img.name,0,0);x=img.x+y[0];y=img.y+y[1];}else{y=f25(img);x=y[0];y=y[1];}return [x,y];}function f26(img){return [pri(img.width),pri(img.height)];}function f25(img){var xy=[img.offsetLeft,img.offsetTop];var ce=img.offsetParent;while(ce!=null){xy[0]+=ce.offsetLeft;xy[1]+=ce.offsetTop;ce=ce.offsetParent;}return xy;}function f18(d,img){var tmp;if(d.images[img]) return d.images[img];for(var i=0;i<d.layers.length;i++){tmp=f18(d.layers[i].document,img);if(tmp) return tmp;}return null;}function f8(d,img,ox,oy){var i;var tmp;if(d.left){ox+=d.left;oy+=d.top;}if(d.document.images[img]) return [ox,oy];for(i=0;i<d.document.layers.length;i++){tmp=f8(d.document.layers[i],img,ox,oy);if(tmp) return [tmp[0],tmp[1]];}return null;}function f0(e){if(IE){f4(e);IsContext=true;}else if(e.which==3){IsContext=true;mX=e.x;mY=e.y;}if(IsContext){f31();cFrame.f19(cntxMenu,mX-1,mY-1,false);return false;}return true;}function f9(){nOM=0;if(!SM) onerror=f14;if(!mFrame) mFrame=cFrame;if(typeof(mFrame)=="undefined"||(NS&&(++NSDelay<2))) window.setTimeout("f9()",10);else{IsFrames=(cFrame!=mFrame);if(NS){mFrame.captureEvents(Event.MOUSEMOVE);mFrame.onmousemove=f24;if(cntxMenu!=""){mFrame.window.captureEvents(Event.MOUSEDOWN);mFrame.window.onmousedown=f0;}nsOWH=f1();window.onresize=rHnd;f5();}if(IE){document.onmousemove=f24;mFrame.document.onmousemove=document.onmousemove;if(cntxMenu) mFrame.document.oncontextmenu=f0;if(IsFrames) mFrame.window.onunload=new Function("mFrame=null;f9()");cFrame.lt=[0,0];}MenusReady=true;}return true;}function f14(sMsg,sUrl,sLine){if(sMsg.substr(0,16)!="Access is denied"&&sMsg!="Permission denied"&&sMsg.indexOf("cursor")==-1) alert("Java Script Error\n"+      "\nDescription:"+sMsg+      "\nSource:"+sUrl+      "\nLine:"+sLine);return true;}function f35(v,s,r,k){var n;if(nOM==0||k==1) n=(v+s>r)?r-s:v;else n=(v+s>r)?pri(om[nOM].style.left)-s:v;return (n<0)?0:n;}function f7(s){if(IsWin||!NS) return s;for(var i=54;i>1;i--) if(s.indexOf("point-size="+i)!=-1) s=xrep(s,"point-size="+i,"point-size="+(i+3));return s;}function f16(t,f){if(!f) cFrame.f=cFrame;if(eval(t)) eval("cFrame.f.clearTimeout("+t+");"+t+"=0");}function xrep(s,f,n){if(s) s=s.split(f).join(n);return s;}function rHnd(){var nsCWH=f1();if((nsCWH[0]!=nsOWH[0])||(nsCWH[1]!=nsOWH[1])) frames["top"].location.reload();}function f34(oName,f){var obj=null;if(!f) f=mFrame;if(NS) obj=f.document.layers[oName];else{obj=(BV>=5?f.document.getElementById(oName):f.document.all[oName]);if(obj) if(obj.id!=oName) obj=null;}return obj;}function gpid(o){return (BV>=5?o.parentNode.parentNode.id:o.parentElement.parentElement.id);}function f5(){for(var l=0;l<mFrame.document.layers.length;l++){var lo=mFrame.document.layers[l];if(lo.layers.length){lo.w=lo.clip.width;lo.h=lo.clip.height;for(var sx=0;sx<lo.layers.length;sx++) for(var sl=0;sl<lo.layers[sx].layers.length;sl++){var slo=mFrame.document.layers[l].layers[sx].layers[sl];if(slo.name.indexOf("EH")>0){slo.document.onmouseup=f21;slo.document.TCode=nTCode[slo.name.split("EH")[1]];}}for(var t=1;t<=tbNum;t++){tb=cFrame.document.layers['dmbTBBack'+t];for(var sl=0;sl<tb.layers['dmbTB'+t].layers.length;sl++){slo=tb.layers['dmbTB'+t].layers[sl];if(slo.name.indexOf('EH')>0){slo.document.onmouseup=f21;slo.document.TCode=nTCode[slo.name.split('EH')[1]];}}}}}}if(NS) with(document){open();write(xrep(f7(f36("grpProducts00117791000804040FFC84022113751001> 1EH122109211003 OnMouseOver=\"cFrame.f13(0,this);status=\'ComputersgrpComputers6  1N22109211002FFC840>84	#8000000>109113>0740Computers281$black_arrow!1010> 1O2210921100280404084	#8FFFFFF>109113>0740Computers281$white_arrow!1010> 2EH2224109211003 OnMouseOver=\"cFrame.f13(0,this);status=\'AccessoriesgrpAccessories6  2N224109211002FFC840>84	#8000000>109113>0740Accessories281$black_arrow!1010> 2O22410921100280404084	#8FFFFFF>109113>0740Accessories281$white_arrow!1010>24610551002>102841 bgcolor=#000000> 4EH4252109211003 OnMouseOver=\"cFrame.f13(0,this);status=\'Catalog Index 4N252109211002FFC840>84	#8000000>109113>0910Catalog Index 4O25210921100280404084	#8FFFFFF>109113>0910Catalog IndexgrpComputers00165861000804040FFC84022161821001> 5EH522157211003 OnMouseOver=\"cFrame.f13(0,this);status=\'High Performance 5N22157211002FFC840>84	#8000000>1013913>01390High Performance 5O2215721100280404084	#8FFFFFF>1013913>01390High Performance 6EH6224157211003 OnMouseOver=\"cFrame.f13(0,this);status=\'Gaming Systems 6N224157211002FFC840>84	#8000000>1013913>01390Gaming Systems 6O22415721100280404084	#8FFFFFF>1013913>01390Gaming Systems 7EH7246157341003 OnMouseOver=\"cFrame.f13(0,this);status=\'Computers for the Home<br>Desktop Systems 7N246157341002FFC840>84	#8000000>1013926>01390Computers for the Home<br>Desktop Systems 7O24615734100280404084	#8FFFFFF>1013926>01390Computers for the Home<br>Desktop SystemsgrpAccessories0090951000804040FFC8402286911001> 8EH82282211003 OnMouseOver=\"cFrame.f13(0,this);status=\'Cables 8N2282211002FFC840>84	#8000000>106413>0640Cables 8O228221100280404084	#8FFFFFF>106413>0640Cables 9EH922482211003 OnMouseOver=\"cFrame.f13(0,this);status=\'Adapters 9N22482211002FFC840>84	#8000000>106413>0640Adapters 9O2248221100280404084	#8FFFFFF>106413>0640Adapters 10EH1024682211003 OnMouseOver=\"cFrame.f13(0,this);status=\'Connectors 10N24682211002FFC840>84	#8000000>106413>0640Connectors 10O2468221100280404084	#8FFFFFF>106413>0640Connectors 11EH1126882211003 OnMouseOver=\"cFrame.f13(0,this);status=\'Batteries 11N26882211002FFC840>84	#8000000>106413>0640Batteries 11O2688221100280404084	#8FFFFFF>106413>0640BatteriesgrpLinks00119511000804040FFC84022115471001> 12EH1222111211003 OnMouseOver=\"cFrame.f13(0,this);status=\'Search the Web 12N22111211002FFC840>84	#8000000>109313>0930Search the Web 12O2211121100280404084	#8FFFFFF>109313>0930Search the Web 13EH13224111211003 OnMouseOver=\"cFrame.f13(0,this);status=\'xFX JumpStart<sup>�</sup> 13N224111211002FFC840>84	#8000000>109313>0930xFX JumpStart<sup>�</sup> 13O22411121100280404084	#8FFFFFF>109313>0930xFX JumpStart<sup>�</sup>")),'%'+'%REL%%',rimPath));close();}f36('');f9();function f36(code){code=xrep(code,""," left=");code=xrep(code,""," top=");code=xrep(code,""," width=");code=xrep(code,""," height=");code=xrep(code,""," z-index=");code=xrep(code,""," visibility=hidden><layer ");code=xrep(code,""," bgColor=#");code=xrep(code,"	","><font face=");code=xrep(code,"","</font>");code=xrep(code,""," point-size=");code=xrep(code,""," color=#");code=xrep(code,""," src=\"");code=xrep(code,"","><div align=left>");code=xrep(code,"","</div>");code=xrep(code,"","\';cFrame.ShowMenu(\'");code=xrep(code,"","<layer");code=xrep(code,"","</layer>");code=xrep(code,"","<ilayer");code=xrep(code,"","</ilayer>");code=xrep(code,"","name=MC");code=xrep(code,"","\';\">");code=xrep(code,""," visibility=hidden>");code=xrep(code,"","\',0,0,true,\'\',");code=xrep(code,"","</b>");code=xrep(code,"","><b");code=xrep(code,"","bgColor=#");code=xrep(code,"","><img");code=xrep(code,""," name=");code=xrep(code," ",");\">");code=xrep(code,"!",".gif\"");code=xrep(code,"#","Tahoma");code=xrep(code,"$","%%REL%%");return code;}function xA(o, n){for(var i=1;i<nTCode.length;i++){if(nTCode[i])nTCode[i]=xrep(nTCode[i],o,n);}}function pri(n){return parseInt(n)}