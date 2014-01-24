<!-- #INCLUDE file="includes/top.asp" -->
<%
page = "Contact"
title = "Contact - Lafayette Testing Services"
metaDescription = "Lafayette Testing Services, Inc. offers high quality non-destructive testing services specializing in radiography, digital radiography, high energy x-ray, magnetic particle, liquid penetrant, and ultrasonics."
metaKeywords = "Lafayette Testing Services contact information, Lafayette Testing Services contact form, conact Lafayette Testing Services, Lafayette Testing Services address, Lafayette Testing Services phone, Lafayette Testing Services location"
%>

<!--METADATA TYPE="typelib" UUID="CD000000-8B95-11D1-82DB-00C04FB1625D" NAME="CDO for Windows Library" -->
<!--METADATA TYPE="typelib" UUID="00000205-0000-0010-8000-00AA006D2EA4" NAME="ADODB Type Library" --> 
<!-- #INCLUDE file="includes/header.asp" -->

<script language="javascript" type="text/javascript">
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

<table width="<%=siteWidth%>" border="0" cellspacing="0" cellpadding="0">
	<tr valign="top">
		<td id="articleContainer">
			<h1>Contact Lafayette Testing Services</h1>
				<p>3710 North Richards Street<br />
				Milwaukee, WI 53212<br />
				T - 1-800-337-4884<br />
				F - 414-332-3652<br />
				<a href="javascript:changeClass('changling_directions','normal')">Driving Directions</a></p>
				<div id="changling_directions" class="visibility_off">
					<p class="noUnderline"><a href="javascript:changeClass('changling_directions','visibility_off')" class="floatRight">[X] Close</a></p><br />
					<div id="map" style="width: 450px; height: 300px"></div><br />
					Enter your starting address:<br />
					<form action="http://maps.google.com/maps" method="get">
						<input type="text" name="saddr" id="saddr" size="35" value="" class="margin_0_10_0_0"/>
				        <input type="submit" class="phone" value="Directions" />
				        <input type="hidden" name="daddr" value="3710 North Richards Street, Milwaukee, WI 53212" />
					    <input type="hidden" name="hl" value="en" />
					</form>
				</div>
				
				<p>Or complete the form below to submit your questions and comments to Lafayette Testing Services.</p>
				<p class="error fineprint">* Indicates a required field.</p>
				<form style="display: inline; margin: 0;" action="/cgi-bin/formmail/formMailer.pl" name="contactForm" method="post" ID="contactForm" onSubmit="return required(this,'First Name, Last Name, Phone, E-Mail Address');">
					<label for="firstName" id="errFirstName"><font class="error">*</font>First Name:</label>
					<input type="text" name="firstName" id="firstName" value="" /><br />
					<label for="lastName" id="errLastName"><font class="error">*</font>Last Name:</label>
					<input type="text" name="lastName" id="lastName" value="" /><br />
					<label for="businessType">Type of Business:</label>
					<input type="text" name="businessType" id="businessType" value="" /><br />
					<label for="address1">Address:</label>
					<input type="text" name="address1" id="address1" value="" /><br />				
					<label for="address2">&nbsp;</label>
					<input type="text" name="address2" id="address2" value="" /><br />
					<label for="city">City:</label>
					<input type="text" name="city" id="city" value="" /><br />
					<label for="state">State/Province:</label>
					<!-- #INCLUDE file="includes/state_pulldown.asp" --><br />
					<label for="country">Country:</label>
					<!-- #INCLUDE file="includes/country_pulldown.asp" --><br />
					<label for="zip">Zip Code:</label>
					<input type="text" class="zip" name="zip" id="zip" value="" maxlength="10" /><br />
					<label for="phone" id="errPhone"><font class="error">*</font>Phone:</label>
					<input type="text" name="phone" id="phone" value="" maxlength="12" /><br />
					<label for="email" id="errEmail"><font class="error">*</font>E-mail Address:</label>
					<input type="text" name="email" id="Text13" value="" /><br />
					<label for="comments">Comments:</label>
					<textarea name="comments" cols="40" rows="8"></textarea><br />
					<input type="submit" class="buttonSubmit" name="Submit" value="Submit" ID="Submit1" /> <input type="reset" class="buttonSubmit" /><br />
				</form>			
			</td>
		<td id="rightRail">
			<img src="<%=images%>about/right_rail_about.jpg" border="0" width="220" height="450" alt="Lafayette Testing Services" /></td>
	</tr>
</table>
<!-- #INCLUDE file="includes/footer.asp" -->