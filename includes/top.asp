
<!DOCTYPE html>
<!--[if IE 8 ]><html lang="en-US" prefix="og: http://ogp.me/ns#" class="no-js ie ie8 lte8 lte9"><![endif]-->
<!--[if IE 9 ]><html lang="en-US" prefix="og: http://ogp.me/ns#" class="no-js ie ie9 lte9"><![endif]-->
<!--[if (gt IE 9)|!(IE)]><!--><html lang="en-US" prefix="og: http://ogp.me/ns#" class="no-js"><!--<![endif]-->

<%
dim developmentLocation
dim sectionPath
dim page
dim address
dim images
dim title
dim metaDescription
dim metaKeywords
dim siteWidth
dim nodeSet

'Form Variables
dim objSendMail, sMessageText

dim navState1, navState2, navState3, navState4, navState5

dim FieldErrorFlag, errFields, form_error

developmentLocation = Request.ServerVariables("HTTP_HOST")
sectionPath = Request.ServerVariables("URL")

Select Case developmentLocation
	Case "localhost:63710"
		address = "http://localhost:63710/"
    Case "www.lafayettetesting.com/review"
		address = "http://www.lafayettetesting.com/review/"
	Case Else
		address = "http://www.lafayettetesting.com/"
End Select

images = address & "images/"

siteWidth="994"

dim selectedContent, image, XMLFile, serviceType

image = Request.QueryString("image")

If Request.QueryString("selectedContent") = "" Then
	 selectedContent = "overview"
Else
	selectedContent = Request.QueryString("selectedContent")
End If

If Request.QueryString("XMLFile") = "" Then
	 XMLFile = "overview"
Else
	XMLFile = Request.QueryString("XMLFile")
End If

serviceType = Request.QueryString("serviceType")

'Form variables
dim firstname, lastname, businessType, phone, email, Address1, Address2, City, company, country, state, zip, fax, comments, error

'Browser/operating system detection
dim UA, OS
dim AnyIE, IE5, IE4, IE3, IE302, AnyNetscape, Netscape7, Netscape6, Netscape5, Netscape4, Netscape3
UA = Request.ServerVariables("HTTP_USER_AGENT")
OS = Request.ServerVariables("HTTP_UA_OS")

'For determining what kind of Browser someone is using

AnyIE = False
IE5 = False
IE4 = False
IE3 = False
IE302 = False
AnyNetscape = False
Netscape7 = False
Netscape6 = False
Netscape5 = False
Netscape4 = False
Netscape3 = False

IF INSTR(UA, "MSIE") THEN
	AnyIE = TRUE
	IF INSTR(UA, "MSIE 5.") THEN
		IE5 = TRUE
	ELSEIF INSTR(UA, "MSIE 4.") THEN
		IE4 = TRUE
	ELSEIF INSTR(UA, "MSIE 3.") THEN
		IE3 = TRUE
		IF INSTR(UA, "MSIE 3.02") THEN
			IE302 = TRUE
		END IF
	END IF
ELSEIF INSTR(UA, "Mozilla") and INSTR(UA, "compatible") = 0 THEN
	AnyNetscape = TRUE
	IF INSTR(UA, "Mozilla/7") THEN
		Netscape7 = TRUE
	ELSEIF INSTR(UA, "Mozilla/6") THEN
		Netscape6 = TRUE
	ELSEIF INSTR(UA, "Mozilla/5") THEN
		Netscape5 = TRUE
	ELSEIF INSTR(UA, "Mozilla/4") THEN
		Netscape4 = TRUE
	ELSEIF INSTR(UA, "Mozilla/3") THEN
		Netscape3 = TRUE
	ELSEIF INSTR(UA, "Mozilla/8") THEN
		AnyIE = TRUE
	END IF
END IF
%>