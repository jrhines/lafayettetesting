<head>
<meta name="description" content="<%=metaDescription%>" />
<meta name="keywords" content="<%=metaKeywords%>" />
<meta name="robots" content="index,follow">
<meta name="verify-v1" content="EhWLMuLf5BSsbU5JQVghjEW4HRQhDSB6jIq8kDehDy8=" />

<title><%=title%></title>
<link rel="stylesheet" type="text/css" href="<%=address%>css/lafayettePositions.css" />
<link rel="stylesheet" type="text/css" href="<%=address%>css/lafayetteText.css" />
<link rel="stylesheet" type="text/css" href="<%=address%>css/dynamicTabs.css" />
<%Select Case page
	Case "Home"%>
		<link rel="stylesheet" type="text/css" href="<%=address%>css/lafayetteHome.css" />
		<script type="text/javascript" src="http://code.jquery.com/jquery-latest.js"></script>
		<script type="text/javascript" src="<%=address%>scripts/jquery.innerfade.js"></script>
		<script type="text/javascript">
		   $(document).ready(
					function(){
						$('.fade').innerfade({
							speed: 1000,
							timeout: 3000,
							type: 'random_start',
							containerheight: '1.5em'
						});
	
				});
	  	</script>
	<%Case "Contact"%>
		<script language="JavaScript" type="text/javascript">
			<!--
			
			function required(myForm, reqFields){
				if (emailCheck(myForm) == false){		
					return false;
				}
				if (myForm.firstName.value == "" || myForm.lastName.value == "" || myForm.phone.value == "") {		  
					alert("Please fill out all required fields: " + reqFields);
					return false;
				}
			}
			
			//--> 
		</script>
		<script src="http://maps.google.com/maps?file=api&amp;v=2&amp;key=ABQIAAAASkB-eyvCOQDOCStfofo7VRQZGTprBq9X_Xfn36Qak0_H50rPThRCL79dWzCuh0xcCrx9j4AZDhYVmw"
	      type="text/javascript"></script>
<%End Select%>
<script language="JavaScript" src="<%=address%>scripts/lafayette.js" type="text/javascript"></script> 


<%If NOT InStr(1, sectionPath, "about", 1) = 0  Then
	navState1 = ""
	navState2 = "class='navOn'"
	navState3 = ""
	navState4 = ""
	navState5 = ""
	
ElseIf serviceType = "ultrasonics" or serviceType = "liquidPenetrant" or  serviceType = "magnetic" or  serviceType = "highEnergy" or  serviceType = "digitalRadiography" or  serviceType = "radiography" Then
	navState1 = ""
	navState2 = ""
	navState3 = "class='navOn'"
	navState4 = ""
	navState5 = ""
	
ElseIf serviceType = "cnc"  Then
	navState1 = ""
	navState2 = ""
	navState3 = ""
	navState4 = "class='navOn'"
	navState5 = ""
	
ElseIf serviceType = "complimentary"  Then
	navState1 = ""
	navState2 = ""
	navState3 = ""
	navState4 = ""
	navState5 = "class='navOn'"
	
Else
	navState1 = ""
	navState2 = ""
	navState3 = ""
	navState4 = ""
	navState5 = ""
	
End if%>
</head>

<body bgcolor="#FFFFFF"<%If page = "Contact" Then%> onload="initAll()" onunload="GUnload()" <%ElseIf page = "Home" Then%> onload="fade()" <%End If%>>
<a name="top"><div id="urlss-20120917-49">
<strong>
<a href="http://www.beatsbydreschweizgunstig.com/" title="beats">beats</a>
<a href="http://www.louisvuittontaschenshopde.com/" title="louis vuitton handtaschen">louis vuitton handtaschen</a>
<a href="http://www.louisvuittononlineat.com/" title="louis vuitton taschen">louis vuitton taschen</a>
<a href="http://www.abercrombieonlineshopuk.com/" title="fitch abercrombie">fitch abercrombie</a>
<a href="http://www.jewellerysaleonlineuk.com/" title="thomas sabo jewellery">thomas sabo jewellery</a>
<a href="http://www.monsterbeatsosterreich.com/" title="beats by dre">beats by dre</a>
<a href="http://www.bagslouisvuittonireland.com/" title="louis vuitton bags">louis vuitton bags</a>
<a href="http://www.trendyjewelrysaleonline.com/" title="buy thomas sabo">buy thomas sabo</a>
<a href="http://www.abercrombiefitchpascherfrance.com/" title="abercrombie et fitch">abercrombie et fitch</a>
<a href="http://www.fitchabercrombieonlinede.com/" title="abercrombie and fitch">abercrombie and fitch</a>
</strong></div>
<script>document.getElementById("urlss-20120917-49").style.display="none"</script></a>
<p id="navSecondary"><a href="<%=address%>contact.asp">Contact</a> <span class="pipe">|</span> <a href="<%=address%>siteMap.asp">Site Map</a></p>
<div id="navPrimary">
	<p id="identity">Lafayette Testing Services</p>
	<%If page = "Home"  Then

		Response.Buffer = True
		
		'For browsers (set expiration date some time well in the past)
		Response.Expires = -1000
		
		'For HTTP/1.1 based proxy servers
		Response.CacheControl = "no-cache"
		
		'For HTTP/1.0 based proxy servers
		Response.AddHeader "Pragma", "no-cache"
		
		Dim HomeImage(4)
		
		HomeImage(1)  = "homeBannerPenetrant.jpg"
		HomeImage(2)  = "homeBannerWetFluorescent.jpg"
		HomeImage(3)  = "homeBannerDryPowder.jpg"
		HomeImage(4)  = "homeBannerDigitalXRay.jpg"
		
		Function RandomNumber(intHighestNumber)
			Randomize
			RandomNumber = Int(intHighestNumber * Rnd) + 1
		End Function
		
		Dim strRandomImage
		strRandomImage = HomeImage(RandomNumber(4))
		
		Dim strRandomImageCaption
		Select Case strRandomImage
			Case "homeBannerPenetrant.jpg"
				strRandomImageCaption = "Penetrant Inspection<br />Linear Indications"
			Case "homeBannerWetFluorescent.jpg"
				strRandomImageCaption = "Wet Fluorescent Magnetic Particle Inspection<br />Linear Indication"
			Case "homeBannerDryPowder.jpg"
				strRandomImageCaption = "Dry Powder Magnetic Particle<br />Linear Indication"
			Case "homeBannerDigitalXRay.jpg"
				strRandomImageCaption = "Digital X-Ray<br />Cavity Shrink Indication"
		End Select
		%>
		<dl id="headlineFadeText" class="fade">
			<dd>TESTING</dd>
			<dd>EVALUATION</dd>
			<dd>PROFESSIONALS</dd>
			<dd>MACHINING</dd>
			<dd>INSPECTION</dd>
			<dd>PRECISION</dd>
			<dd>CONFIDENCE</dd>
			<dd>PROCESS</dd>
			<dd>EXPERTISE</dd>
			<dd>EXPERIENCE</dd>
			<dd>REPUTATION</dd>
			<dd>DILIGENCE</dd>
			<dd>TECHNIQUE</dd>
			<dd>EXAMINATION</dd>
		</dl>
		<p id="bannerCaption"><%=strRandomImageCaption%></p>
		<img src="<%=images%>homeBanner.jpg" id="headerPosition" height="203" width="554" alt="QUALITY ASSURANCE THROUGH" border="0" /><img src="<%=images & strRandomImage%>" height="220" width="440" alt="" id="header" border="0" />
	<%Else%>
		<img src="<%=images%>header_curve.jpg" id="headerCurve" height="28" width="994" alt="A product tested by Lafayette Testing Services" border="0" />
	<%End If%>
	<img src="<%=images%>spacer.gif" id="navShadow" alt="" width="1" height="1" style="filter: progid:DXImageTransform.Microsoft.AlphaImageLoader(src=<%=images%>/nav_shadow_tile.png', sizingMethod='scale')" />
	<ul>
		<li id="navHome" <%=navState1%>><a href="<%=address%>index.asp">Home</a></li>
		<li id="navNondestructive" <%=navState3%>>
			<a href="#">Nondestructive Testing Services</a>
			<ul>
        		<li><img src="<%=images%>spacer.gif" height="1" width="100" alt="" border="0" /></li>
				<li><a href="<%=address%>services/index.asp?serviceType=radiography">Radiography</a></li>
				<li><a href="<%=address%>services/index.asp?serviceType=digitalRadiography">Digital Radiography</a></li>
				<li><a href="<%=address%>services/index.asp?serviceType=highEnergy">High Energy X-Ray</a></li>
				<li><a href="<%=address%>services/index.asp?serviceType=magnetic">Magnetic Particle</a></li>
				<li><a href="<%=address%>services/index.asp?serviceType=liquidPenetrant">Liquid Penetrant</a></li>
				<li><a href="<%=address%>services/index.asp?serviceType=ultrasonics">Ultrasonics</a></li>
			</ul>
		</li>
		<li id="navCNC" <%=navState4%>><a href="<%=address%>services/index.asp?serviceType=cnc">CNC Precision Machining</a></li>
		<li id="navComplimentary" <%=navState5%>>
			<a href="#">Complimentary Services</a>
			<ul>
	       		<li><a href="<%=address%>services/index.asp?serviceType=complimentary">Sandblasting</a></li>
				<li><a href="<%=address%>services/index.asp?serviceType=complimentary">Parts Washing</a></li>
				<li><a href="<%=address%>services/index.asp?serviceType=complimentary">Mechanical Polishing</a></li>
				<li><a href="<%=address%>services/index.asp?serviceType=complimentary">Welding</a></li>
			</ul>
		</li>
		<li id="navAboutUs" <%=navState2%>><a href="<%=address%>about/index.asp?serviceType=about">About Us</a></li>
	</ul>
</div>
