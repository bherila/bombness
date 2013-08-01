SI_footer.clear			= true;
SI_footer.container		= 'content';
SI_footer.minHeight		= 410;
SI_footer.extendShallow = false;
SI_footer.bottomOut		= true;
	
window.onload = function() {
	SI_menu('nav-main'); // ,'nav-sub'
	SI_clearFooter();
	SI_initializeTabs();
	SI_initializeGroups();
	SI_initializeSwapImg();
	if (SI_footer.clear) {
		window.onresize = SI_clearFooter;
		SI.resize.initialize();
		}
	};