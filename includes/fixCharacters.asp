<%
	dim returnString, returnStringNonXSL
	
	' this will replace problematic characters in the XML with "safe" equivalent
	Function fixCharacters( inStr, loc )
		returnString = inStr
		
		if loc = "preTransform" then

			' the initial character in the following line is a literal n-dash symbol
			returnString = Replace(returnString, "�", "-")
			returnString = Replace(returnString, "&#8211;", "-")
			returnString = Replace(returnString, "&ndash;", "-")
			returnString = Replace(returnString, "&#8212;", "--")
			returnString = Replace(returnString, "&mdash;", "--")
			
			returnString = Replace(returnString, "'", "&apos;")
			returnString = Replace(returnString, "&rsquo;", "&apos;")
			returnString = Replace(returnString, "&#8217;", "&apos;")
			returnString = Replace(returnString, "�", "&apos;")
			' the following is different from the above
			returnString = Replace(returnString, "�", "&apos;")
			returnString = Replace(returnString, "�", "&apos;")
			
			'opening curly quotes - middle double quote is literal symbol
			returnString = Replace(returnString, "�", "&#34;")
			'closing curly quotes - middle double quote is literal symbol
			returnString = Replace(returnString, "�", "&#34;")

			returnString = Replace(returnString, "&nbsp;", " ")
			
			returnString = Replace(returnString, "�", "&#189;")
			returnString = Replace(returnString, "�", "&#169;")
			returnString = Replace(returnString, "�", "&#174;")
			returnString = Replace(returnString, "&trade;", "trademark153_1")
			returnString = Replace(returnString, "�", "trademark153_2")
			
			'n with tilde
			returnString = Replace(returnString, "�", "&#241;")
			returnString = Replace(returnString, "&ntilde;", "&#241;")
			returnString = Replace(returnString, "(ntilde)", "&#241;")
			'a, A acute
			returnString = Replace(returnString, "�", "&#225;")
			returnString = Replace(returnString, "�", "&#193;")
			returnString = Replace(returnString, "&aacute;", "&#225;")
			returnString = Replace(returnString, "&Aacute;", "&#193;")
			returnString = Replace(returnString, "(aacute)", "&#225;")
			returnString = Replace(returnString, "(Aacute)", "&#193;")
			'a, A grave
			returnString = Replace(returnString, "�", "&#224;")
			returnString = Replace(returnString, "�", "&#192;")
			returnString = Replace(returnString, "&agrave;", "&#224;")
			returnString = Replace(returnString, "&Agrave;", "&#192;")
			returnString = Replace(returnString, "(agrave)", "&#224;")
			returnString = Replace(returnString, "(Agrave)", "&#192;")
			'e, E acute
			returnString = Replace(returnString, "�", "&#233;")
			returnString = Replace(returnString, "�", "&#201;")
			returnString = Replace(returnString, "&eacute;", "&#233;")
			returnString = Replace(returnString, "&Eacute;", "&#201;")
			returnString = Replace(returnString, "(eacute)", "&#233;")
			returnString = Replace(returnString, "(Eacute)", "&#201;")
			'e, E grave
			returnString = Replace(returnString, "�", "&#232;")
			returnString = Replace(returnString, "�", "&#200;")
			returnString = Replace(returnString, "&egrave;", "&#232;")
			returnString = Replace(returnString, "&Egrave;", "&#200;")
			returnString = Replace(returnString, "(egrave)", "&#232;")
			returnString = Replace(returnString, "(Egrave)", "&#200;")
			'i, I acute
			returnString = Replace(returnString, "�", "&#237;")
			returnString = Replace(returnString, "�", "&#205;")
			returnString = Replace(returnString, "&iacute;", "&#237;")
			returnString = Replace(returnString, "&Iacute;", "&#205;")
			returnString = Replace(returnString, "(iacute)", "&#237;")
			returnString = Replace(returnString, "(Iacute)", "&#205;")
			'o, O acute
			returnString = Replace(returnString, "�", "&#243;")
			returnString = Replace(returnString, "�", "&#211;")
			returnString = Replace(returnString, "&oacute;", "&#243;")
			returnString = Replace(returnString, "&Oacute;", "&#211;")
			returnString = Replace(returnString, "(oacute)", "&#243;")
			returnString = Replace(returnString, "(Oacute)", "&#211;")
			'u, U acute
			returnString = Replace(returnString, "�", "&#250;")
			returnString = Replace(returnString, "�", "&#218;")
			returnString = Replace(returnString, "&uacute;", "&#250;")
			returnString = Replace(returnString, "&Uacute;", "&#218;")
			returnString = Replace(returnString, "(uacute)", "&#250;")
			returnString = Replace(returnString, "(Uacute)", "&#218;")
			
			returnString = Replace(returnString, "(C)", "&#169;")
			returnString = Replace(returnString, "(c)", "&#169;")
			returnString = Replace(returnString, "(R)", "&#174;")
			returnString = Replace(returnString, "(r)", "&#174;")
			
			'ellipses
			returnString = Replace(returnString, "&hellip;", "&#133;")
			
		elseif loc = "postTransform" then
		
			'n with tilde
			returnString = Replace(returnString, "�", "&#241;")
			returnString = Replace(returnString, "&ntilde;", "&#241;")
			'a, A acute
			returnString = Replace(returnString, "�", "&#225;")
			returnString = Replace(returnString, "�", "&#193;")
			returnString = Replace(returnString, "&aacute;", "&#225;")
			returnString = Replace(returnString, "&Aacute;", "&#193;")
			'a, A grave
			returnString = Replace(returnString, "�", "&#224;")
			returnString = Replace(returnString, "�", "&#192;")
			returnString = Replace(returnString, "&agrave;", "&#224;")
			returnString = Replace(returnString, "&Agrave;", "&#192;")
			'e, E acute
			returnString = Replace(returnString, "�", "&#233;")
			returnString = Replace(returnString, "�", "&#201;")
			returnString = Replace(returnString, "&eacute;", "&#233;")
			returnString = Replace(returnString, "&Eacute;", "&#201;")
			'e, E grave
			returnString = Replace(returnString, "�", "&#232;")
			returnString = Replace(returnString, "�", "&#200;")
			returnString = Replace(returnString, "&egrave;", "&#232;")
			returnString = Replace(returnString, "&Egrave;", "&#200;")
			'i, I acute
			returnString = Replace(returnString, "�", "&#237;")
			returnString = Replace(returnString, "�", "&#205;")
			returnString = Replace(returnString, "&iacute;", "&#237;")
			returnString = Replace(returnString, "&Iacute;", "&#205;")
			'o, O acute
			returnString = Replace(returnString, "�", "&#243;")
			returnString = Replace(returnString, "�", "&#211;")
			returnString = Replace(returnString, "&oacute;", "&#243;")
			returnString = Replace(returnString, "&Oacute;", "&#211;")
			'u, U acute
			returnString = Replace(returnString, "�", "&#250;")
			returnString = Replace(returnString, "�", "&#218;")
			returnString = Replace(returnString, "&uacute;", "&#250;")
			returnString = Replace(returnString, "&Uacute;", "&#218;")
		
			returnString = Replace(returnString, "�", "&#176;")
			returnString = Replace(returnString, "�", "&#149;")
			returnString = Replace(returnString, "�", "&#189;")
			returnString = Replace(returnString, "�", "&#169;")
			returnString = Replace(returnString, "�", "&#174;")
			'returnString = Replace(returnString, "�", "&#153;")
			returnString = Replace(returnString, "trademark153_1", "&#153;")
			returnString = Replace(returnString, "trademark153_2", "&#153;")
			
			returnString = Replace(returnString, "&amp;quot;", "&#34;")
			returnString = Replace(returnString, "&apos;", "&#39;")
			
			'ellipses
			returnString = Replace(returnString, "&hellip;", "&#133;")
			
			
		end if
		
		fixCharacters = returnString
	End Function
	
	'this will replace problematic characters from the database with "safe" equivalent
	'for instance, copyright and registered symbols display as question marks under
	'UTF-8 encoding
	Function fixCharactersNonXSL( inStr )
	
		returnStringNonXSL = inStr
		
		' the initial character in the following line is a literal n-dash symbol
		returnStringNonXSL = Replace(returnStringNonXSL, "�", "-")
		returnStringNonXSL = Replace(returnStringNonXSL, "&#8211;", "-")
		returnStringNonXSL = Replace(returnStringNonXSL, "&ndash;", "-")
		returnStringNonXSL = Replace(returnStringNonXSL, "&#8212;", "--")
		returnStringNonXSL = Replace(returnStringNonXSL, "&mdash;", "--")
		
		returnStringNonXSL = Replace(returnStringNonXSL, "&rsquo;", "&#39;")
		returnStringNonXSL = Replace(returnStringNonXSL, "&#8217;", "&#39;")
		returnStringNonXSL = Replace(returnStringNonXSL, "�", "&#39;")
		' the following is different from the above
		returnStringNonXSL = Replace(returnStringNonXSL, "�", "&#39;")
		returnStringNonXSL = Replace(returnStringNonXSL, "�", "&#39;")
		
		'opening curly quotes - middle double quote is literal symbol
		returnStringNonXSL = Replace(returnStringNonXSL, "�", "&#34;")
		'closing curly quotes - middle double quote is literal symbol
		returnStringNonXSL = Replace(returnStringNonXSL, "�", "&#34;")
		
		returnStringNonXSL = Replace(returnStringNonXSL, "&#160;", "&nbsp;")
		
		'degrees
		returnStringNonXSL = Replace(returnStringNonXSL, "�", "&#176;")
		returnStringNonXSL = Replace(returnStringNonXSL, "�", "&#149;")
		returnStringNonXSL = Replace(returnStringNonXSL, "�", "&#189;")
		
		returnStringNonXSL = Replace(returnStringNonXSL, "�", "&#169;")
		returnStringNonXSL = Replace(returnStringNonXSL, "�", "&#174;")
		returnStringNonXSL = Replace(returnStringNonXSL, "&trade;", "&#153;")
		returnStringNonXSL = Replace(returnStringNonXSL, "�", "&#153;")
		
		returnStringNonXSL = Replace(returnStringNonXSL, "(C)", "&#169;")
		returnStringNonXSL = Replace(returnStringNonXSL, "(c)", "&#169;")
		returnStringNonXSL = Replace(returnStringNonXSL, "(R)", "&#174;")
		returnStringNonXSL = Replace(returnStringNonXSL, "(r)", "&#174;")
		
		'n with tilde
		returnStringNonXSL = Replace(returnStringNonXSL, "�", "&#241;")
		'a, A acute
		returnStringNonXSL = Replace(returnStringNonXSL, "�", "&#225;")
		returnStringNonXSL = Replace(returnStringNonXSL, "�", "&#193;")
		'a, A grave
		returnStringNonXSL = Replace(returnStringNonXSL, "�", "&#224;")
		returnStringNonXSL = Replace(returnStringNonXSL, "�", "&#192;")
		'e, E acute
		returnStringNonXSL = Replace(returnStringNonXSL, "�", "&#233;")
		returnStringNonXSL = Replace(returnStringNonXSL, "�", "&#201;")
		'e, E grave
		returnStringNonXSL = Replace(returnStringNonXSL, "�", "&#232;")
		returnStringNonXSL = Replace(returnStringNonXSL, "�", "&#200;")
		'� Circumflex '07/27/07 MTM
		returnStringNonXSL = Replace(returnStringNonXSL, "�", "&#234;")
		'i, I acute
		returnStringNonXSL = Replace(returnStringNonXSL, "�", "&#237;")
		returnStringNonXSL = Replace(returnStringNonXSL, "�", "&#205;")
		'o, O acute
		returnStringNonXSL = Replace(returnStringNonXSL, "�", "&#243;")
		returnStringNonXSL = Replace(returnStringNonXSL, "�", "&#211;")
		'u, U acute
		returnStringNonXSL = Replace(returnStringNonXSL, "�", "&#250;")
		returnStringNonXSL = Replace(returnStringNonXSL, "�", "&#218;")
		
		'ellipses
		returnString = Replace(returnString, "&hellip;", "&#133;")
	
		fixCharactersNonXSL = returnStringNonXSL
		
	End Function
	
%>
