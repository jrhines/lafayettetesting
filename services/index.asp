<!-- #Include File="../includes/top.asp" -->
<!-- #Include File="../includes/fixCharacters.asp" -->
<!-- #Include File="../includes/XMLDataController.asp" -->
<!-- #Include File="../includes/dynamicTabs.asp" -->
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
	metaKeywords = oXMLDC_1.getXMLData("single", "/Testing/Service[@name='" & serviceType & "']/metaKeywords", "")

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

<!-- #INCLUDE file="../includes/header.asp" -->
<!-- #Include File="../includes/FlashWrapper.asp" -->

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
			
			<div class="floatRight">					
				<!-- #INCLUDE file="../includes/addThis.asp" -->
			</div></td>
		<td id="rightRail">
			<script language="javascript" type="text/javascript">
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
<!-- #INCLUDE file="../includes/footer.asp" -->