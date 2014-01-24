<%
	dim returnString, returnStringNonXSL
	
	' this will replace problematic characters in the XML with "safe" equivalent
	Function fixCharacters( inStr, loc )
		returnString = inStr
		
		if loc = "preTransform" then

			' the initial character in the following line is a literal n-dash symbol
			returnString = Replace(returnString, "–", "-")
			returnString = Replace(returnString, "&#8211;", "-")
			returnString = Replace(returnString, "&ndash;", "-")
			returnString = Replace(returnString, "&#8212;", "--")
			returnString = Replace(returnString, "&mdash;", "--")
			
			returnString = Replace(returnString, "'", "&apos;")
			returnString = Replace(returnString, "&rsquo;", "&apos;")
			returnString = Replace(returnString, "&#8217;", "&apos;")
			returnString = Replace(returnString, "’", "&apos;")
			' the following is different from the above
			returnString = Replace(returnString, "’", "&apos;")
			returnString = Replace(returnString, "‘", "&apos;")
			
			'opening curly quotes - middle double quote is literal symbol
			returnString = Replace(returnString, "“", "&#34;")
			'closing curly quotes - middle double quote is literal symbol
			returnString = Replace(returnString, "”", "&#34;")

			returnString = Replace(returnString, "&nbsp;", " ")
			
			returnString = Replace(returnString, "½", "&#189;")
			returnString = Replace(returnString, "©", "&#169;")
			returnString = Replace(returnString, "®", "&#174;")
			returnString = Replace(returnString, "&trade;", "trademark153_1")
			returnString = Replace(returnString, "™", "trademark153_2")
			
			'n with tilde
			returnString = Replace(returnString, "ñ", "&#241;")
			returnString = Replace(returnString, "&ntilde;", "&#241;")
			returnString = Replace(returnString, "(ntilde)", "&#241;")
			'a, A acute
			returnString = Replace(returnString, "á", "&#225;")
			returnString = Replace(returnString, "Á", "&#193;")
			returnString = Replace(returnString, "&aacute;", "&#225;")
			returnString = Replace(returnString, "&Aacute;", "&#193;")
			returnString = Replace(returnString, "(aacute)", "&#225;")
			returnString = Replace(returnString, "(Aacute)", "&#193;")
			'a, A grave
			returnString = Replace(returnString, "à", "&#224;")
			returnString = Replace(returnString, "À", "&#192;")
			returnString = Replace(returnString, "&agrave;", "&#224;")
			returnString = Replace(returnString, "&Agrave;", "&#192;")
			returnString = Replace(returnString, "(agrave)", "&#224;")
			returnString = Replace(returnString, "(Agrave)", "&#192;")
			'e, E acute
			returnString = Replace(returnString, "é", "&#233;")
			returnString = Replace(returnString, "É", "&#201;")
			returnString = Replace(returnString, "&eacute;", "&#233;")
			returnString = Replace(returnString, "&Eacute;", "&#201;")
			returnString = Replace(returnString, "(eacute)", "&#233;")
			returnString = Replace(returnString, "(Eacute)", "&#201;")
			'e, E grave
			returnString = Replace(returnString, "è", "&#232;")
			returnString = Replace(returnString, "È", "&#200;")
			returnString = Replace(returnString, "&egrave;", "&#232;")
			returnString = Replace(returnString, "&Egrave;", "&#200;")
			returnString = Replace(returnString, "(egrave)", "&#232;")
			returnString = Replace(returnString, "(Egrave)", "&#200;")
			'i, I acute
			returnString = Replace(returnString, "í", "&#237;")
			returnString = Replace(returnString, "Í", "&#205;")
			returnString = Replace(returnString, "&iacute;", "&#237;")
			returnString = Replace(returnString, "&Iacute;", "&#205;")
			returnString = Replace(returnString, "(iacute)", "&#237;")
			returnString = Replace(returnString, "(Iacute)", "&#205;")
			'o, O acute
			returnString = Replace(returnString, "ó", "&#243;")
			returnString = Replace(returnString, "Ó", "&#211;")
			returnString = Replace(returnString, "&oacute;", "&#243;")
			returnString = Replace(returnString, "&Oacute;", "&#211;")
			returnString = Replace(returnString, "(oacute)", "&#243;")
			returnString = Replace(returnString, "(Oacute)", "&#211;")
			'u, U acute
			returnString = Replace(returnString, "ú", "&#250;")
			returnString = Replace(returnString, "Ú", "&#218;")
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
			returnString = Replace(returnString, "ñ", "&#241;")
			returnString = Replace(returnString, "&ntilde;", "&#241;")
			'a, A acute
			returnString = Replace(returnString, "á", "&#225;")
			returnString = Replace(returnString, "Á", "&#193;")
			returnString = Replace(returnString, "&aacute;", "&#225;")
			returnString = Replace(returnString, "&Aacute;", "&#193;")
			'a, A grave
			returnString = Replace(returnString, "à", "&#224;")
			returnString = Replace(returnString, "À", "&#192;")
			returnString = Replace(returnString, "&agrave;", "&#224;")
			returnString = Replace(returnString, "&Agrave;", "&#192;")
			'e, E acute
			returnString = Replace(returnString, "é", "&#233;")
			returnString = Replace(returnString, "É", "&#201;")
			returnString = Replace(returnString, "&eacute;", "&#233;")
			returnString = Replace(returnString, "&Eacute;", "&#201;")
			'e, E grave
			returnString = Replace(returnString, "è", "&#232;")
			returnString = Replace(returnString, "È", "&#200;")
			returnString = Replace(returnString, "&egrave;", "&#232;")
			returnString = Replace(returnString, "&Egrave;", "&#200;")
			'i, I acute
			returnString = Replace(returnString, "í", "&#237;")
			returnString = Replace(returnString, "Í", "&#205;")
			returnString = Replace(returnString, "&iacute;", "&#237;")
			returnString = Replace(returnString, "&Iacute;", "&#205;")
			'o, O acute
			returnString = Replace(returnString, "ó", "&#243;")
			returnString = Replace(returnString, "Ó", "&#211;")
			returnString = Replace(returnString, "&oacute;", "&#243;")
			returnString = Replace(returnString, "&Oacute;", "&#211;")
			'u, U acute
			returnString = Replace(returnString, "ú", "&#250;")
			returnString = Replace(returnString, "Ú", "&#218;")
			returnString = Replace(returnString, "&uacute;", "&#250;")
			returnString = Replace(returnString, "&Uacute;", "&#218;")
		
			returnString = Replace(returnString, "°", "&#176;")
			returnString = Replace(returnString, "•", "&#149;")
			returnString = Replace(returnString, "½", "&#189;")
			returnString = Replace(returnString, "©", "&#169;")
			returnString = Replace(returnString, "®", "&#174;")
			'returnString = Replace(returnString, "™", "&#153;")
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
		returnStringNonXSL = Replace(returnStringNonXSL, "–", "-")
		returnStringNonXSL = Replace(returnStringNonXSL, "&#8211;", "-")
		returnStringNonXSL = Replace(returnStringNonXSL, "&ndash;", "-")
		returnStringNonXSL = Replace(returnStringNonXSL, "&#8212;", "--")
		returnStringNonXSL = Replace(returnStringNonXSL, "&mdash;", "--")
		
		returnStringNonXSL = Replace(returnStringNonXSL, "&rsquo;", "&#39;")
		returnStringNonXSL = Replace(returnStringNonXSL, "&#8217;", "&#39;")
		returnStringNonXSL = Replace(returnStringNonXSL, "’", "&#39;")
		' the following is different from the above
		returnStringNonXSL = Replace(returnStringNonXSL, "’", "&#39;")
		returnStringNonXSL = Replace(returnStringNonXSL, "‘", "&#39;")
		
		'opening curly quotes - middle double quote is literal symbol
		returnStringNonXSL = Replace(returnStringNonXSL, "“", "&#34;")
		'closing curly quotes - middle double quote is literal symbol
		returnStringNonXSL = Replace(returnStringNonXSL, "”", "&#34;")
		
		returnStringNonXSL = Replace(returnStringNonXSL, "&#160;", "&nbsp;")
		
		'degrees
		returnStringNonXSL = Replace(returnStringNonXSL, "°", "&#176;")
		returnStringNonXSL = Replace(returnStringNonXSL, "•", "&#149;")
		returnStringNonXSL = Replace(returnStringNonXSL, "½", "&#189;")
		
		returnStringNonXSL = Replace(returnStringNonXSL, "©", "&#169;")
		returnStringNonXSL = Replace(returnStringNonXSL, "®", "&#174;")
		returnStringNonXSL = Replace(returnStringNonXSL, "&trade;", "&#153;")
		returnStringNonXSL = Replace(returnStringNonXSL, "™", "&#153;")
		
		returnStringNonXSL = Replace(returnStringNonXSL, "(C)", "&#169;")
		returnStringNonXSL = Replace(returnStringNonXSL, "(c)", "&#169;")
		returnStringNonXSL = Replace(returnStringNonXSL, "(R)", "&#174;")
		returnStringNonXSL = Replace(returnStringNonXSL, "(r)", "&#174;")
		
		'n with tilde
		returnStringNonXSL = Replace(returnStringNonXSL, "ñ", "&#241;")
		'a, A acute
		returnStringNonXSL = Replace(returnStringNonXSL, "á", "&#225;")
		returnStringNonXSL = Replace(returnStringNonXSL, "Á", "&#193;")
		'a, A grave
		returnStringNonXSL = Replace(returnStringNonXSL, "à", "&#224;")
		returnStringNonXSL = Replace(returnStringNonXSL, "À", "&#192;")
		'e, E acute
		returnStringNonXSL = Replace(returnStringNonXSL, "é", "&#233;")
		returnStringNonXSL = Replace(returnStringNonXSL, "É", "&#201;")
		'e, E grave
		returnStringNonXSL = Replace(returnStringNonXSL, "è", "&#232;")
		returnStringNonXSL = Replace(returnStringNonXSL, "È", "&#200;")
		'ê Circumflex '07/27/07 MTM
		returnStringNonXSL = Replace(returnStringNonXSL, "ê", "&#234;")
		'i, I acute
		returnStringNonXSL = Replace(returnStringNonXSL, "í", "&#237;")
		returnStringNonXSL = Replace(returnStringNonXSL, "Í", "&#205;")
		'o, O acute
		returnStringNonXSL = Replace(returnStringNonXSL, "ó", "&#243;")
		returnStringNonXSL = Replace(returnStringNonXSL, "Ó", "&#211;")
		'u, U acute
		returnStringNonXSL = Replace(returnStringNonXSL, "ú", "&#250;")
		returnStringNonXSL = Replace(returnStringNonXSL, "Ú", "&#218;")
		
		'ellipses
		returnString = Replace(returnString, "&hellip;", "&#133;")
	
		fixCharactersNonXSL = returnStringNonXSL
		
	End Function
	
%>
