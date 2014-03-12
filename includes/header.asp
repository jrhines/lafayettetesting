<head>
<meta name="description" content="<%=metaDescription%>" />
<meta name="robots" content="index,follow" />
<meta name="verify-v1" content="EhWLMuLf5BSsbU5JQVghjEW4HRQhDSB6jIq8kDehDy8=" />

<title><%=title%></title>
<link rel="stylesheet" href="<%=address%>css/lafayettePositions.css" />
<link rel="stylesheet" href="<%=address%>css/lafayetteText.css" />
<link rel="stylesheet" href="<%=address%>css/dynamicTabs.css" />
<%Select Case page
	Case "Home"%>
		<link rel="stylesheet" href="<%=address%>css/lafayetteHome.css" />
		<script src="http://code.jquery.com/jquery-latest.js"></script>
		<script src="<%=address%>scripts/jquery.innerfade.js"></script>
		<script>
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
		<script>
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
		<script src="http://maps.google.com/maps?file=api&amp;v=2&amp;key=ABQIAAAASkB-eyvCOQDOCStfofo7VRQZGTprBq9X_Xfn36Qak0_H50rPThRCL79dWzCuh0xcCrx9j4AZDhYVmw"></script>
<%End Select%>
<script src="<%=address%>scripts/lafayette.js"></script> 


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

<body<%If page = "Contact" Then%> onload="initAll()" onunload="GUnload()"<%ElseIf page = "Home" Then%> onload="fade()"<%End If%>>
<p id="navSecondary"><a href="<%=address%>contact.asp">Contact</a> <span class="pipe">|</span> <a href="<%=address%>siteMap.asp">Site Map</a></p>
<div id="navPrimary">
	<a id="identity" href="<%=address%>index.asp"><img src="<%=images%>ltsLogoHeader.gif" alt="Lafayette Testing Services" /><span>Lafayette Testing Services</span></a>
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
            <dt></dt>
			<dd>testing</dd>
			<dd>evaluation</dd>
			<dd>professionals</dd>
			<dd>machining</dd>
			<dd>inspection</dd>
			<dd>precision</dd>
			<dd>confidence</dd>
			<dd>process</dd>
			<dd>expertise</dd>
			<dd>experience</dd>
			<dd>reputation</dd>
			<dd>diligence</dd>
			<dd>technique</dd>
			<dd>examination</dd>
		</dl>
		<p id="bannerCaption"><%=strRandomImageCaption%></p><br />
		<img src="<%=images%>homeBanner.jpg" alt="Quality assurance through" /><img src="<%=images & strRandomImage%>" id="header" alt="" />
	<%Else%>
		<img src="<%=images%>header_curve.jpg" id="headerCurve" alt="A product tested by Lafayette Testing Services" />
	<%End If%>
	<img src="<%=images%>spacer.gif" id="navShadow" alt="" style="filter: progid:DXImageTransform.Microsoft.AlphaImageLoader(src=<%=images%>/nav_shadow_tile.png', sizingMethod='scale')" />
	<ul>
		<li id="navHome" <%=navState1%>><a href="<%=address%>index.asp">Home</a></li>
		<li id="navNondestructive" <%=navState3%>>
			<a href="#">Nondestructive Testing Services</a>
			<ul>
        		<li><img src="<%=images%>spacer.gif" height="1" width="100" alt="" /></li>
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