function FlashWrapper(){}

FlashWrapper.prototype.initFlashWrapper = function(thisMovie, thisWidth, thisHeight, thisWmode, thisStaticImg, thisStaticImgLink, thisStaticImgAlt, thisStaticImgMap)
{

	this.thisMovie = unescape(thisMovie);
	this.thisWidth = thisWidth;
	this.thisHeight = thisHeight;
	this.thisWmode = thisWmode;
	this.thisStaticImg = unescape(thisStaticImg);
	this.thisStaticImgLink = unescape(thisStaticImgLink);
	this.thisStaticImgAlt = unescape(thisStaticImgAlt);
	this.thisStaticImgMap = thisStaticImgMap;
	
	// Check to see that Flash is installed and that it is the correct version
	// If not, show default image

	this.flashinstalled = 0;
	this.flashversion = 0;
	this.MSDetect = "false";
	
	if ( navigator.plugins && navigator.plugins.length )
	{
		x = navigator.plugins["Shockwave Flash"];
		
		if ( x )
		{
			flashinstalled = 2;
			
			if ( x.description )
			{
				y = x.description;
				flashversion = y.charAt(y.indexOf('.')-1);
			}
		}
		else
		{
			flashinstalled = 1;
		}
		
		if ( navigator.plugins["Shockwave Flash 2.0"] )
		{
			flashinstalled = 2;
			flashversion = 2;
		}
		if( flashinstalled == "2" && flashversion >= "7" )
		{
			this.writeFlashObj();
		}
		else {				
			document.write("")
		}
	}
	else if ( navigator.mimeTypes && navigator.mimeTypes.length )
	{
		x = navigator.mimeTypes['application/x-shockwave-flash'];
		
		if ( x && x.enabledPlugin )
		{
			flashinstalled = 2;
		}
		else 
		{
			flashinstalled = 1;
		}
		if ( flashinstalled == "2" && flashversion >= "7" )
		{
			this.writeFlashObj();
		}
		else
		{
			this.writeStaticImg();
		}
	}
	else
	{
		MSDetect = "true";
		//---------------------------------------------------------------------------------------------------------------------
		// 	Check the browser...we're looking for ie/win
		
		var isIE  = ( navigator.appVersion.indexOf("MSIE") != -1 ) ? true : false;    // true if we're on ie
		var isWin = ( navigator.appVersion.toLowerCase().indexOf("win") != -1 ) ? true : false; // true if we're on windows
		
		if( isIE && isWin )
		{
			// these variables are defined in FlashWrapperIEDetectVarsInit.js and set for IE
			// in FlashWrapperIEDetectVarsInit.vbs
		 	if ((flash6Installed == true) || (flash7Installed == true) || (flash8Installed == true) || (flash9Installed == true)) 
			{
				this.writeFlashObj();
			}
			else {
				this.writeStaticImg();
			}
		}
	}
}	


FlashWrapper.prototype.writeFlashObj = function()
{
	var objTag;

	if ( this.thisWmode == "" )
	{
		this.thisWmode = "false";
	}
	
	objTag = '<OBJECT type="application/x-shockwave-flash" data="' + this.thisMovie + '" '
			+ 'WIDTH="' + this.thisWidth + '" HEIGHT="' + this.thisHeight + '">'
			+ '<PARAM NAME="MOVIE" VALUE="' + this.thisMovie + '">'
			+ '<PARAM NAME="PLAY" VALUE="true">'
			+ '<PARAM NAME="QUALITY" VALUE="high">'
			+ '<PARAM NAME="MENU" VALUE="false">'
			+ '<PARAM NAME="WMODE" VALUE="' + this.thisWmode + '">'							
			+ '<\/OBJECT>';

	document.write(objTag);
}	
		
			
FlashWrapper.prototype.writeStaticImg = function()
{
	if ( this.thisStaticImg != "" ) 
	{
	
		if( this.thisStaticImgMap == "" )
		{
			if ( this.thisStaticImgLink != "" )
			{
				document.write("<a href=\"" + this.thisStaticImgLink + "\">");
			}
			
			document.write("<img src=\"" + this.thisStaticImg + "\" alt=\"" + this.thisStaticImgAlt + "\">");
			
			if ( this.thisStaticImgLink != "" )
			{
				document.write("<\/a>");
			}
		}
		
		else
		{
			document.write("<img src=\"" + this.thisStaticImg + "\" alt=\"" + this.thisStaticImgAlt + "\" usemap=\"#" + this.thisStaticImgMap + "\">");
		}

	}
	
}
