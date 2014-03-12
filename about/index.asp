<!-- #Include Virtual="/includes/top.asp" -->
<!-- #Include Virtual="/includes/fixCharacters.asp" -->
<!-- #Include Virtual="/includes/XMLDataController.asp" -->
<!-- #Include Virtual="/includes/dynamicTabs.asp" -->
<%
page = "About Us"
title = "About Lafayette Testing Services, Inc. (LTS) - quality - nondestructive testing - CNC precision machining - customer service"
metaDescription = "Lafayette Testing Services delivers high quality nondestructive testing services at competitive prices using the most updated technology to achieve accurate results within the shortest possible time frame."

dim oXMLDC_1, newAryParams(2,1)
	
	newAryParams(0,0) = "selectedContent"
	newAryParams(0,1) = selectedContent
	newAryParams(1,0) = "XMLFile"
	newAryParams(1,1) = XMLFile
	newAryParams(2,0) = "image"
	newAryParams(2,1) = image
	
	
	set oXMLDC_1 = new XMLDataController
	oXMLDC_1.XMLFile = address & "about/" & XMLFile & ".xml"
	oXMLDC_1.initXMLDataController
	oXMLDC_1.aryParams = newAryParams
%>

<!-- #Include Virtual="/includes/header.asp" -->
<!-- #Include Virtual="/includes/FlashWrapper.asp" -->
<%'Dynamic Tabs V1.0

dim strTabTitles, strTabSelected, strTabType, strQSExclude, oDynamicTabs

strTabTitles = "overview|lab_diversification|image_gallery"

strTabSelected = Request("selectedContent")

If strTabSelected = "" Then
	strTabSelected = "overview"		
End If

strTabType = "round"

'strQSExclude = "dest"

set oDynamicTabs = new dynamicTabs

oDynamicTabs.tabSelected = strTabSelected
oDynamicTabs.tabTitles = strTabTitles
oDynamicTabs.tabType = strTabType
oDynamicTabs.QSExclude = strQSExclude
%>
<table id="content-table">
	<tr>
		<td id="articleContainer">
			<%response.write(oDynamicTabs.getTabs())%>
			
			<%= oXMLDC_1.getXMLData("complete", "", "aboutContent.xsl") %>
			<div class="floatRight">					
				<!-- #Include Virtual="/includes/addThis.asp" -->
			</div></td>
		<td id="rightRail">
			<%If selectedContent="image_gallery" Then%>
				<script>
					var flashObj1 = new FlashWrapper();
					var flashURL = escape("../images/about/right_rail_dedicated_professionals.swf");
					var flashWidth = "220";
					var flashHeight = "474";
					var flashWmode = "transparent";
					var noFlashImgURL = escape("../images/about/right_rail_about.jpg");
					var noFlashImgHREF = escape("");
					var noFlashImgAltText = escape("<%=serviceType%>");
					var imgMapName = "";
			
					// arguments  flash file :: width :: height :: wmode setting :: static img file :: static img url :: static img alt text :: img map
					flashObj1.initFlashWrapper(flashURL, flashWidth, flashHeight, flashWmode, noFlashImgURL, noFlashImgHREF, noFlashImgAltText, imgMapName);
				</script>	
			<%Else%>
				<img src="<%=images%>about/right_rail_about.jpg" alt="Lafayette Testing Services" />
			<%End If%>
		</td>
	</tr>
</table>
<!-- #Include Virtual="/includes/footer.asp" -->