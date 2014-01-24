<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'	INPUTS:
'	public property XMLFile --> xml file name --> file to transform
'
'	public property DOMObject --> dom object name --> object to transform
'
'   public property aryParams --> 	2-dimensional array containing params to pass 
'									into xsl p_aryParams(x,0) is key and 
'									p_aryParams(x,1) is value 
'
'	OUTPUT:
'	result of Public Function getXMLData, as HTML
'
'	AUTHOR: Chip Kreis 8/2/07
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

class XMLDataController

	private p_XMLFile, p_DOMObj, p_aryParams, p_XMLFileError, p_objXML
	
	'''''''''''''''''''''''''''''''''''''''''''''
		public property let XMLFile(value)
			p_XMLFile = value
		end property
	
	'''''''''''''''''''''''''''''''''''''''''''''
		public property let DOMObj(value)
			p_DOMObj = value
		end property
	
	'''''''''''''''''''''''''''''''''''''''''''''
		public property let aryParams(value)
			p_aryParams = value
		end property
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		public sub initXMLDataController
		
			'check to see if xml file exists
			'if dom object exists load it
			'if error, set p_xmlFileError to true
			'create xml object
			
			set p_objXML = Server.CreateObject("Microsoft.XMLDOM")
						
			If p_DOMObj <> "" Then			
				p_objXML.loadXML(p_DOMobj)
			Else
				p_objXML.async = false
				p_objXML.setProperty "ServerHTTPRequest", true
				p_objXML.load(p_XMLFile)
				
				if p_objXML.parseError.errorcode <> 0 then
					p_XMLFileError = true
				else
					p_XMLFileError = false
				end if
			End If
			
		end sub
		
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		public function getXMLData(inType, inXPath, inXSL)
		
			dim thisReturnHTML, thisOutputHTML, thisStrXML
			dim thisNode, thisNodeList, thisTransformNode, thisTransformNodeValue
			dim thisXSLT, thisXSLProc, thisObjXSL, thisObjParams
			dim thisObjRoot, thisObjXMLPartial, thisObjXMLPartialStr
		
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			' inType: represents unit of xml data being returned
			' 		- "single" 		- 	value of single node, or non-list set 
			'							of child nodes
			'		- "multiple" 	- 	node list [XSL optional]
			'		- "complete" 	- 	xml document element transformed by 
			'							XSL file specified [XSL required]
			'		- "contained" 	- 	variation of "complete" - path to
			'							XSL file contained in xml node "Transform"
			'							So if the name of the xml file is contained
			'							in the query string, it (the xml file) can
			'							tell us what file needs to be used in its
			'							transformation
			'
			' inXPath: xpath to desired data
			'		- only an option with "single" and "multiple"
			'
			' inXSL: XSL file for transform
			'
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			
			if p_xmlFileError <> "true" then
			
				select case inType
			
					case "single"
					
						'get single node from the xml object
						set thisNode = p_objXML.documentElement.selectSingleNode(inXPath)
						
						if inXSL <> "" then
							
							set thisObjXMLPartial = Server.CreateObject("Microsoft.XMLDOM")
							thisObjXMLPartial.async = false
							
							thisObjXMLPartialStr = "<root>" & thisNode.xml & "</root>"
							thisObjXMLPartial.loadXML(thisObjXMLPartialStr)
							set thisObjRoot = thisObjXMLPartial.documentElement
							
							thisReturnHTML =  transformXML(thisObjXMLPartial, inXSL)
				
						else
							thisReturnHTML = thisNode.text
						end if
						
					case "multiple"
					
						thisObjXMLPartialStr = ""
						
						Set thisNodeList = p_objXML.documentElement.selectNodes(inXPath)
				
						if inXSL <> "" then
							
							set thisObjXMLPartial = Server.CreateObject("Microsoft.XMLDOM")
							thisObjXMLPartial.async = false
							
							For Each thisNode In thisNodeList
								thisObjXMLPartialStr = thisObjXMLPartialStr + thisNode.xml
							Next
							
							thisObjXMLPartialStr = "<root>" & thisObjXMLPartialStr & "</root>"
							thisObjXMLPartial.loadXML(thisObjXMLPartialStr)
			
							thisReturnHTML =  transformXML(thisObjXMLPartial, inXSL)
				
						end if
						
					case "complete"
					
						thisReturnHTML = transformXML(p_objXML, inXSL)
						
					case "contained"
						
						'get xsl file from xml "Transform" node
						set thisTransformNode = p_objXML.documentElement.selectSingleNode("Transform")
						thisTransformNodeValue = thisTransformNode.text
					
						if thisTransformNodeValue <> "" then
							thisReturnHTML = transformXML(p_objXML, thisTransformNodeValue)
						else
							thisReturnHTML = "<b>No file name was provided for xsl transform.</b>"
						end if
						
					
					case default
						thisReturnHTML = "<b>No data return type specified.</b>"
				
				end select
				
			else
			
				thisReturnHTML = "<b>An error occurred while loading the xml.</b>"
				
			end if
			
			getXMLData = thisReturnHTML
		
		end function
		
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		private function transformXML(inObjXML, inXSL)
		
			dim thisReturnHTML
			dim thisOutputHTML
			dim thisObjRoot
			dim thisObjXSL
			dim thisXSLT
			dim thisXSLProc
			dim thisParamIdx
			dim thisXMLStr
			
			'pre-transform character fix
			thisXMLStr = inObjXML.xml
			thisXMLStr = fixCharacters( thisXMLStr, "preTransform" )
			inObjXML.loadXML(thisXMLStr)
			
			'DEBUG
			'THIS IS NOT RETURNING ANYTHING transformXML = "inObjXML.xml: " & inObjXML.xml
			'THIS IS OK - transformXML = "inXSL: " & inXSL
			'exit function
			
			set thisObjRoot = inObjXML.documentElement
			set thisObjXSL = Server.CreateObject( "Msxml2.FreeThreadedDOMDocument" )
			thisObjXSL.async = false
			
			if inXSL <> "" then
				thisObjXSL.load(Server.Mappath(inXSL))
			
				if (thisObjXSL.parseError.errorCode = 0) then
						
					if not thisObjRoot Is nothing then
						
						Set thisXSLT = Server.CreateObject("Msxml2.xsltemplate")					
						thisXSLT.stylesheet = thisObjXSL
						
						Set thisXSLProc = thisXSLT.createProcessor()
						
						'ADD PARAMETERS IF THEY EXIST
						if isArray(p_aryParams) then
				
							for thisParamIdx = 0 to ubound(p_aryParams, 1)
								thisXSLProc.addParameter p_aryParams(thisParamIdx,0), cstr(p_aryParams(thisParamIdx,1))
							next 
							
						end if
						
						thisXSLProc.input = inObjXML
						thisXSLProc.transform()
						thisOutputHTML = thisXSLProc.output
						
						'post-transform character fix
						thisOutputHTML = fixCharacters( thisOutputHTML, "postTransform" )
						
						Set thisXSLProc = Nothing
						Set thisXSLT = Nothing
						
						'set return value of transform to variable thisOutputHTML
						thisReturnHTML = thisOutputHTML
						
					else
						thisReturnHTML = "<b>An error occurred during the xsl transformation.</b>"
					end if
					
				end if
			
			else
				thisReturnHTML = "<b>No file name was provided for xsl transform.</b>"
			end if
			
			transformXML = thisReturnHTML
			
		end function
	
end class

%>