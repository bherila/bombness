/* DOM check v1.0
 * Confirms that the browser is capable of utilizing most SI_functions
 * and hides unnecessary elements from these browsers. Apply the .non-dom class
 * to any element that is made redundant due to scripting. 
 * ie. go buttons next to select menus with onchange event handlers
 */
if (SI_dom()) { document.write('<style type="text/css">.non-dom { display: none !important; .only-dom {display: inherit !important; }<'+'/style>\n'); }

/* SI_dom() v1.0
 * Confirms that the browser is capable of 
 * getElementById and getElementsByTagName
 */
function SI_dom() { return new Boolean(document.getElementById && document.getElementsByTagName); };

/* SI_submitForm() v1.1 */
function SI_submitForm(f) {
	if (!SI_dom()) return true;
	f.submit();
	}

/* SI_windowOpen() v1.0
 * Accessible pop-up windows that load the href in the current window
 * if JavaScript is disabled. 
 */
var SI_openWindow = new Object();
function SI_windowOpen(elem,center,w,h,opt) {
	if (!SI_dom()) return true;
	
	var src	= elem.href;
	var wnm	= (elem.target=='')?'SI_schoolSuite':((elem.target.indexOf('_')==0)?'SI'+elem.target:elem.target);
	var opt = (opt=='')?',scrollbars=yes,resizable=yes':((opt.indexOf(',')!=0)?','+opt:opt);
	var ctr	= (center)?',left='+((screen.availWidth-w)/2)+',top='+((screen.availHeight-h)/2):'';
	
	// A window has already been opened
	if (!SI_openWindow.closed && SI_openWindow.name==wnm && SI_openWindow.location) {
		SI_openWindow.location = src;
		}
	else {
		SI_openWindow = window.open(src,wnm,'width='+w+',height='+h+ctr+opt);
		if (!SI_openWindow.opener) SI_openWindow.opener = self;
		}
	if (window.focus) SI_openWindow.focus();
	return false;
	};

/* SI_clearFooter() v1.3 Customized for Altamont School
 * Use this function when absolutely positioning section navigation, 
 * primarary and secondary content and the footer falls below any one
 * of them. Absolute positioning avoids the unwanted wrapping that
 * sometimes occurs in IE PC when floats are used for layout.
 */
var SI_footer=new Object();
/* Configuration Defaults */
SI_footer.clear			= false;
SI_footer.container		= 'content';	// The inner container div
SI_footer.minHeight		= 410;	// Minimum page height in pixels
SI_footer.extendShallow = true;	// Extends shallow children so that repeating backgrounds 
								// will repeat all the way to the footer
SI_footer.bottomOut		= true;	// Whether the footer should just clear the contents of the 
								// inner-container div or snap to the bottom of the window 
								// when content is too short
function SI_clearFooter() {
	var d = document,w=window,dE=d.documentElement,dB=d.body,h;
	if (!SI_footer.clear || !d.getElementById || !dB.offsetHeight) return;
	
	var ic = d.getElementById(SI_footer.container); // The inner container div
	if (!ic) return; // Added 4.08.23
	//ic.style.height = '0';		// Resets the height
	var oh = [];		// Holds the offset heights of each child div
	var icTop = ic.offsetTop;
	var winHeight	= (typeof(w.innerHeight)=='number')?w.innerHeight:(dE&&dE.clientHeight)?dE.clientHeight:(dB&&dB.clientHeight)?dB.clientHeight:0;
	var footHeight	= d.getElementById('footer').offsetHeight;
	oh[0] = SI_footer.minHeight-icTop-footHeight;	// Minimum height of the inner-container div

	// Make an array of the rendered heights of the contained elements
	for (var i=0;i<ic.childNodes.length;i++) {
		c = ic.childNodes[i];
		if (c.nodeName=='DIV') {
			//alert(c.id + ': ' + c.offsetHeight + ' (debug)');
			if (c.id=='inner-content') {
				c.style.height = SI_getContainedHeight(c)+'px';
				oh[oh.length] = c.offsetHeight+c.offsetTop;
				}
			else if (c.id=='nav-tert') { continue; }
			else {
				c.style.height = 'auto';
				oh[oh.length] = c.offsetHeight+c.offsetTop;
				}
			}
		}
	// Determine the max height
	h = 0; for (var k=0;k<oh.length;k++) { h = (oh[k]>h)?oh[k]:h;}
	// Force footer to the bottom of the window 
	if (SI_footer.bottomOut) { h = ((icTop+h+footHeight) < winHeight)?winHeight-footHeight-icTop:h; }
	// Extend shorter elements
	if (SI_footer.extendShallow) { for (var i=0;i<ic.childNodes.length;i++) { var c = ic.childNodes[i]; if (c.nodeName=='DIV' && c.id!='nav-tert') { c.style.height = h+'px'; }}}
	ic.style.height = h+'px';
	};

function SI_getContainedHeight(e) {
	var oh = [];
	for (var i=0;i<e.childNodes.length;i++) { var c = e.childNodes[i]; if (c.nodeName=='DIV') { c.style.height = 'auto'; oh[oh.length] = c.offsetHeight+c.offsetTop; }}
	// Determine the max height
	var h = 0; for (var k=0;k<oh.length;k++) { h = (oh[k]>h)?oh[k]:h; }
	// alert(h);
	return h;
	}

/* SI_menu() v2.0
 * 
 * Holds IEs hand through what can be accomplished using only
 * CSS in other modern browsers. That's right, I said "modern."
 * In a medium that moves as fast the internet you didn't really
 * consider a 4 year old browser like IE modern did you? :P
 * 
 * Includes nodeName clarification which corrects
 * potential problems in IE 5.x (both platforms)
 * 
 * Now correctly applies the hover className to the li and not the link.
 * Removed RegExp from onmouseout by precaching the class names which speeds 
 * up the script phenomenally.
 */
function SI_menu() {
	var d = document;
	var isSafari 	= (navigator.userAgent.indexOf('Safari') != -1);
	var isIE 		= (navigator.appName == "Microsoft Internet Explorer");
	var isWinIE		= (isIE && window.print);
	
	/* Safari doesn't need any help...
	 * Gecko needs a little help remembering which parent elements are currently :hover
	 * I have no idea what to do with Opera...
	 * IE needs the most help...
	 */
	if (!SI_dom() || window.opera || isSafari) return;
	
	var m=SI_menu.arguments;

	for(i=0; i<m.length; i++) {
		if (!d.getElementById(m[i])) continue; // Added 4.08.23
		for (var l=0; (lnk=d.getElementById(m[i]).getElementsByTagName("a")[l]); l++) {
			// If there are any nested menus...
			if (lnk.parentNode.childNodes.length > 1) {
				li = lnk.parentNode; // The containing <li>
				for (var n=0; n < li.childNodes.length; n++) {
					node = li.childNodes[n];
					if (node.nodeName=="UL") {
						li.ul = node; // The sibling <ul> (submenu)
						delete node;
						
						li.classDefault		= li.className;
						li.classHover		= li.className+((li.className=='')?'hover':' hover');
						
						li.isIE				= isIE;
						li.isWinIE			= isWinIE;
						li.onmouseover		= SI_showMenu;
						li.onmouseout		= SI_hideMenu;
						}
					}
				}
			}
		}
	}
function SI_showMenu() {
	// SI_debug();
	this.className = this.classHover;
	if (this.isIE) {
		this.style.zIndex = 100;
		if (this.isWinIE) SI_toggleSelects('hidden'); 
		}
	}
function SI_hideMenu() {
	this.className = this.classDefault;
	if (this.isIE) {
		this.style.zIndex = 1;
		if (this.isWinIE) SI_toggleSelects('visible'); 
		}
	}
function SI_toggleSelects(state) {
	var d = document;
	for (var i=0; (sel=d.getElementsByTagName('select')[i]); i++) {
		sel.style.visibility = state;
		}
	}
/* SI_debug() v2.0
 * Sets the onmouseover event for all elements to display the 
 * current element's ancestor tree in the window status bar.
 */
function SI_debug() {
	var d = document;
	if (!d.getElementsByTagName) return;
	
	var all = (d.all)?d.all:d.getElementsByTagName('*');
	for (i=0; i<all.length;i++) {
		//var oldmouseover = (all[i].onmouseover)?all[i].onmouseover:function(){};
		all[i].onmouseover = function(e) {
			//oldmouseover();
			if (!e) var e = window.event;
			var status = '';
			var done = false;
			var ths = this;
			while (!done) {
				status = ths.nodeName.toLowerCase()+((ths.className!='')?'.'+ths.className.replace(/ /g,'.'):'')+((ths.id!='')?'#'+ths.id:'')+((status!='')?' > ':'')+status;
				done = (ths.nodeName=='HTML')?true:false;
				if (!done) ths=ths.parentNode;
				}
			
			this.status = status;
			window.status = ((this.status.length>128)?'...':'')+this.status.substr(this.status.length-124);
			e.cancelBubble = true;
			e.returnValue = false;
			if (e.stopPropagation) e.stopPropagation();
			return true;
			}
		}
	}

/* SI_initializeTabs()/SI_activateTab() v1.0
 * Accessible content tabbing that requires only one hit to the server
 * and degrades to simple HTML anchors when JavaScript is disabled.
 * Need to add: Support for multiple classes on ul.tabs li
 *
 * Changing `className` instead of `style.display` would allow for more 
 * control when dealing with print stylesheets.
 */
var SI_tabs=new Object();
function SI_initializeTabs() {
	for (var tab in SI_tabs) { 
		SI_activateTab(tab,SI_tabs[tab].active);
		}
	}
function SI_activateTab(tabGroup,activeTab) {
	var d = document;
	if (!d.getElementById) return;
	
	for (i=0; i<SI_tabs[tabGroup].tabs.length; i++) {
		//alert('tab-'+tabGroup+'-'+SI_tabs[tabGroup].tabs[i])
		tab = 'tab-'+tabGroup+'-'+SI_tabs[tabGroup].tabs[i];
		d.getElementById(tab).className = '';
		d.getElementById(tabGroup+'-'+SI_tabs[tabGroup].tabs[i]).style.display = 'none';
		}
	d.getElementById(tabGroup+'-'+activeTab).style.display = 'block';
	d.getElementById('tab-'+tabGroup+'-'+activeTab).className = 'active-tab';
	
	// Redraw the footer...
	SI_clearFooter();
	};

/* SI_initializeGroups()/SI_toggleGroups() v1.0
 * 
 */
var SI_groups=new Object();
function SI_initializeGroups() {
	var d = document;
	if (!d.getElementById) return;
	for (var group in SI_groups) {
		d.getElementById(group+'-toggle').style.display = 'inline';
		for (i=0; i<SI_groups[group].items.length; i++) {
			anItem = SI_groups[group].items[i];
			d.getElementById(anItem+'-toggle').style.display = 'inline';
			}
		SI_groups[group].expanded = SI_groups[group].items.length;
		SI_toggleGroups(group,'',d.getElementById(group+'-toggle'));
		}
	};
function SI_toggleGroups(group,item,toggle) {
	var d = document;
	if (!d.getElementById) return;
	
	var state = (toggle.innerHTML.toLowerCase().indexOf('hide ')!=-1);
	var display = (state)?'none':'block';
	var action = (state)?'Show ':'Hide ';
	if (item!='') {
		d.getElementById(item).style.display = display;
		toggle.innerHTML = action+SI_groups[group].label;
		SI_groups[group].expanded = (state)?SI_groups[group].expanded-1:SI_groups[group].expanded+1;
		d.getElementById(group+'-toggle').innerHTML = (SI_groups[group].expanded == SI_groups[group].items.length)?'Hide All':'Show All';
		}
	else {
		for (i=0; i<SI_groups[group].items.length; i++) {
			item = SI_groups[group].items[i];
			d.getElementById(item+'-toggle').innerHTML = ((state)?'Show ':'Hide ')+SI_groups[group].label;
			d.getElementById(item).style.display = display;
			SI_groups[group].expanded = (state)?0:SI_groups[group].items.length;
			}
		toggle.innerHTML = action+'All';
		}
	
	// Redraw the footer...
	SI_clearFooter();
	};


/* SI_galleryRedraw() and related v1.4
 * Client-side image gallery. An array of images and a configuration
 * object need to be defined in the HTML.
 *
 * 1.4	: Moves "x 0f y" to its own unique container
 */
var SI_gallery=new Object(); var SI_imgs=new Array(); SI_imgs[0]=null;
function SI_galleryRedraw() {
	if (!SI_dom()) return false;
	var d = document;
	var g = SI_gallery;
		
	// Update the image, caption, count and description
	imgSrc  = SI_imgs[g.imgActive][1]+'thumb/'+g.imgPrefix+SI_imgs[g.imgActive][0];
	imgLnk  = 'http://www.delbarton.org/manager/PublicFileView.asp?myurl='+SI_imgs[g.imgActive][1]+SI_imgs[g.imgActive][0];
	imgLnk += '&Title='+SI_imgs[g.imgActive][2];
	imgLnk += '&Content='+SI_imgs[g.imgActive][3];
	
	img = d.getElementById('SI_galleryImg');
	img.onload = SI_clearFooter;
	img.src=imgSrc;
	// img.alt=SI_imgs[g.imgActive][3].replace(/(<[^>]*>)/g,'');
	img.parentNode.href=imgLnk;
	img.parentNode.onclick=function() {
		SI_windowOpen(this,true,SI_imgs[g.imgActive][4],SI_imgs[g.imgActive][5],'scrollbars=yes,resizable=yes'); return false;
		}
	d.getElementById('SI_galleryImgTitle').innerHTML=SI_imgs[g.imgActive][2]; //+' <span><em>'+g.imgActive+'<'+'/em> of '+g.imgTotal+'<'+'/span>';
	d.getElementById('SI_galleryImgNumOf').innerHTML='<strong>'+g.imgActive+'<'+'/strong> of '+g.imgTotal;
	// Should alread be contained by a `<p>`
	d.getElementById('SI_galleryImgDesc').innerHTML=SI_imgs[g.imgActive][3];
	
	// Hide or show the previous/next buttons as appropriate
	d.getElementById('SI_galleryImgPrev').innerHTML=(g.imgActive==1)?'<em>Previous Image</em>':'<a href="#Previous Image" onclick="SI_galleryImgPrev(); return false;">Previous Image</a>';
	d.getElementById('SI_galleryImgNext').innerHTML=(g.imgActive==g.imgTotal)?'<em>Next Image</em>':'<a href="#Next Image" onclick="SI_galleryImgNext(); return false;">Next Image</a>';
	
	// Redraw with the current sets thumbnails 
	var t = d.getElementById('SI_galleryThumbs');
	t = '';
	for (i=g.setImg; i<g.setImg+g.thumbTotal;i++) {
		if (i<=g.imgTotal) {
			var className = (i==g.setImg)?'first-child':((i==g.setImg+g.thumbTotal)?'last-child':'');
			if (i==g.imgActive) { className += (className=='')?'active':' active'; }
			className = (className!='')?' class="'+className+'"':'';
			t += '<li'+className+'><a href="#image'+i+'" onclick="SI_galleryImgSelect('+i+'); return false;"><img src="'+SI_imgs[i][1]+'thumb/'+g.thumbPrefix+SI_imgs[i][0]+'" alt="'+SI_imgs[i][2]+'" border="0" /><'+'/a><'+'/li>\n';
			}
		}
	d.getElementById('SI_galleryThumbs').innerHTML = t;
	
	var sets = '<strong>Set:</strong> ';
	for (i=1; i<=g.setTotal; i++) {
		if (i==g.setActive) {
			sets += i+' ';
			}
		else {
			sets += '<a href="#Load Set '+i+'" onclick="SI_gallerySetSelect('+i+'); return false;">'+i+'<'+'/a> ';
			}
		if (i!=g.setTotal) sets += '<i>|<'+'/i> ';
		}
	d.getElementById('SI_gallerySets').innerHTML = sets;
	}
// Displays the selected image
function SI_galleryImgSelect(imgNo) {
	if (SI_gallery.imgActive == imgNo) return;
	SI_gallery.imgActive = imgNo;
	SI_galleryRedraw();
	}
// Displays the next image
function SI_galleryImgNext() {
	var g = SI_gallery;
	if (g.imgActive<g.imgTotal) {
		g.imgActive++;
		if (g.imgActive >= g.setImg+g.thumbTotal) {
			g.setImg = g.imgActive;
			g.setActive++;
			}
		SI_galleryRedraw();
		}
	}
// Displays the previous image
function SI_galleryImgPrev() {
	var g = SI_gallery;
	if (g.imgActive>1) {
		g.imgActive--;
		if (g.imgActive < g.setImg) {
			g.setActive--;
			g.setImg = (g.setActive*g.thumbTotal)-(g.thumbTotal-1);
			}
		SI_galleryRedraw();
		}
	}
// Displays the selected set of images
function SI_gallerySetSelect(setNo) {
	var g = SI_gallery;
	g.setActive = setNo;
	g.setImg = (g.setActive*g.thumbTotal)-(g.thumbTotal-1);
	g.imgActive = g.setImg;
	SI_galleryRedraw();
	}
// Displays the previous set of images
function SI_gallerySetPrev() {
	var g = SI_gallery;
	if (g.setActive!=1) {
		g.setActive--;
		g.setImg = (g.setActive*g.thumbTotal)-(g.thumbTotal-1);
		g.imgActive = g.setImg;
		SI_galleryRedraw();
		}
	}
// Displays the next set of images
function SI_gallerySetNext() {
	var g = SI_gallery;
	if (g.setActive<g.setTotal) {
		g.setActive++;
		g.setImg = (g.setActive*g.thumbTotal)-(g.thumbTotal-1);
		g.imgActive = g.setImg;
		SI_galleryRedraw();
		}
	}

// v1.2
function SI_initializeSwapImg() {
	var d = document;
	var SI_preloadImgs = new Array();
	for (i=0;img=d.images[i];i++) {	
		if (img.src.indexOf('over=')!=-1) {
			img.defaultsrc	= img.src;
			img.oversrc		= img.src.replace(/^(.+)over=/i,'');
			img.onmouseover	= function () { this.src = this.oversrc; };
			img.onmouseout	= function () { this.src = this.defaultsrc; };
			
			SI_preloadImgs[i] = new Image();
			SI_preloadImgs[i].src = img.oversrc;
			}
		}
	}


/* SI and SI.resize Objects v1.0
 * Resize creates a hidden div and monitors changes in it's offsetHeight
 * to determine when the user resizes text.
 */
var SI		= new Object();
SI.resize	= new Object();
SI.resize.initialize = function() {
	if (document.all && window.print) { return; }
	
	var c = document.createElement('div');
	
	c.style.position	= 'fixed';
	c.style.top			= '0';
	c.style.visibility	= 'hidden';
	c.style.width		= '10em';
	c.style.height		= '10em';
	
	SI.resize.control = document.body.appendChild(c);
	SI.resize.h = 0;
	window.setInterval('SI.resize.detectChange()',50);
	}
SI.resize.detectChange = function() {
	var o = SI.resize.h;
	SI.resize.h = SI.resize.control.offsetHeight;
	if (o!=SI.resize.h) SI.resize.hasOccurred();
	}
/* SI.resize.hasOccurred() v1.0
 * Add functions that you want to be called whenever 
 * the user resizes text to this method.
 */
SI.resize.hasOccurred = function() {
	SI_clearFooter();
	}

function SI_quicklinks(e) {
	var url=e.value;
	if (url!='') { window.location=url; }
	else { e.selectedIndex=0; }
	}


/* SI_customizeButton() v1.0
 * Used to toggle customize options in calendar block view
 */
SI_customizeButtonState = 'open';
function SI_customizeButton() {
	if (!SI_dom()) return;
	var btn = document.getElementById('customize-btn');
	var frm = document.getElementById('customize');
	if (SI_customizeButtonState=='open') {
		// btn.src ='/images/calendars/block/btn_customize_open.gif';
		SI_customizeButtonState = 'close';
		frm.style.display = 'block';
		}
	else {
		// btn.src ='/images/calendars/block/btn_customize.gif';
		SI_customizeButtonState = 'open';
		frm.style.display = 'none';
		}
	}