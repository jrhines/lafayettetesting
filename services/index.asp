<!-- #Include Virtual="/includes/top.asp" -->
<!-- #Include Virtual="/includes/fixCharacters.asp" -->
<!-- #Include Virtual="/includes/XMLDataController.asp" -->
<!-- #Include Virtual="/includes/dynamicTabs.asp" -->
<%
	set oXMLDC_1 = new XMLDataController
	oXMLDC_1.XMLFile = address & "services/metaTags.xml"
	oXMLDC_1.initXMLDataController
	
	dim oXMLDC_2
	set oXMLDC_2 = new XMLDataController
	oXMLDC_2.XMLFile = address & "services/overview.xml"
	oXMLDC_2.initXMLDataController
	
	title = oXMLDC_2.getXMLData("single", "/Testing/Service[@name='" & serviceType & "']/Headline", "") & " - Lafayette Testing Services"
	metaDescription = oXMLDC_1.getXMLData("single", "/Testing/Service[@name='" & serviceType & "']/metaDescription", "")

	dim oXMLDC_1, newAryParams(3,1)
	
	newAryParams(0,0) = "serviceType"
	newAryParams(0,1) = serviceType
	newAryParams(1,0) = "selectedContent"
	newAryParams(1,1) = selectedContent
	newAryParams(2,0) = "XMLFile"
	newAryParams(2,1) = XMLFile
	newAryParams(3,0) = "image"
	newAryParams(3,1) = image
	
	
	set oXMLDC_1 = new XMLDataController
	oXMLDC_1.XMLFile = address & "services/" & XMLFile & ".xml"
	oXMLDC_1.initXMLDataController
	oXMLDC_1.aryParams = newAryParams
%>

<!-- #Include Virtual="/includes/header.asp" -->
<!-- #Include Virtual="/includes/FlashWrapper.asp" -->

<%'Dynamic Tabs V1.0

dim strTabTitles, strTabSelected, strTabType, strQSExclude, oDynamicTabs

strTabTitles = "overview|image_gallery"

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
<table border="0" width="<%=siteWidth%>" cellspacing="0" cellpadding="0">
	<tr valign="top">
		<td id="articleContainer">
			<%response.write(oDynamicTabs.getTabs())%>
	
			<%= oXMLDC_1.getXMLData("complete", "", "content_disciplines.xsl") %>
			
            <div id="service-information">
                <ul class="resp-tabs-list">
		            <li>Technical Specifications</li>
		            <li>Advantages</li>
				    <li>Applications</li>
		        </ul> 
		        <div class="resp-tabs-container">                                                        
		            <div>
                        <h3>Technical Specifications</h3>
                        <p>who</p>
                    </div>
                    <div>
	            	    <h3>Advantages</h3>
					    where
		            </div>
		            <div>
		                <h3>Applications</h3>
					    what
		            </div>
                </div>
            </div>

			<div class="floatRight">					
				<!-- #Include Virtual="/includes/addThis.asp" -->
			</div></td>
		<td id="rightRail">
			<script>
				var flashObj1 = new FlashWrapper();
				var flashURL = escape("../images/nondestructive/right_rail_<%=serviceType%>.swf");
				var flashWidth = "220";
				var flashHeight = "474";
				var flashWmode = "transparent";
				var noFlashImgURL = escape("../images/nondestructive/right_rail_<%=serviceType%>.jpg");
				var noFlashImgHREF = escape("");
				var noFlashImgAltText = escape("<%=serviceType%>");
				var imgMapName = "";
		
				// arguments  flash file :: width :: height :: wmode setting :: static img file :: static img url :: static img alt text :: img map
				flashObj1.initFlashWrapper(flashURL, flashWidth, flashHeight, flashWmode, noFlashImgURL, noFlashImgHREF, noFlashImgAltText, imgMapName);
			</script>
		</td>
	</tr>
</table>
<!-- #Include Virtual="/includes/footer.asp" -->