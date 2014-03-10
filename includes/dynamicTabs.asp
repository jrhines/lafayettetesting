<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'	INPUTS:
'	public property TabTitles -->   pipe-delimited string containing tab names
'
'	public property TabSelected --> string containing the name of the currently 
'									selected tab
'   public property TabType --> 	string containing either "sqare" or "round".
'									used to determine sqare or rounded corners.
'	public property QSExclude -->	pipe-delimited string containing the names
'									of query string paramaters to exclude. 
'
'	OUTPUT: 
'	result of Public Function GetTabs, as HTML
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Class dynamicTabs

	Private p_tabTitles, p_tabSelected, p_tabType, p_QSExclude

	'''''''''''''''''''''''''''''''''''''''''''''
		Public Property Let tabTitles(value)
			p_tabTitles = value
		End Property
		
	'''''''''''''''''''''''''''''''''''''''''''''
		Public Property Let tabSelected(value)
			p_tabSelected = value
		End Property
	
	'''''''''''''''''''''''''''''''''''''''''''''
		Public Property Let tabType(value)
			p_tabType = value
		End Property
	
	'''''''''''''''''''''''''''''''''''''''''''''
		Public Property Let QSExclude(value)
			p_QSExclude = value
		End Property
	
	'''''''''''''''''''''''''''''''''''''''''''''
	
		Public Function getTabs()
			
			dim p_tabArray, p_tabCount, x, y, p_tabTitlePre, p_tabTitlePost, p_tabState, p_strOutput, p_URL, p_QS
			dim p_QSArray, p_QSCount, p_QSItem, p_QSName, p_QSEqual, p_excludeArray, p_excludeCount, p_excludeItem, p_delimiter		
			
			p_delimiter = "|"
			p_URL = Request.ServerVariables("URL") 
			p_QS = Request.ServerVariables("QUERY_STRING")
			
			' Remove any unwanted query string name/value pairs
			If p_QS <> "" Then				
				p_QS = Replace(p_QS,"&amp;",p_delimiter)						
				p_QSArray = split(p_QS,p_delimiter)
				p_QSCount = uBound(p_QSArray)
				p_QSExclude = p_QSExclude & p_delimiter & "selectedContent" & p_delimiter & "XMLFile"
				p_excludeArray = split(p_QSExclude,p_delimiter)
				p_excludeCount = uBound(p_excludeArray)
				
				x = 0
				Do while x <= p_QSCount
					p_QSItem = p_QSArray(x)
					p_QSEqual = inStr(1,p_QSItem,"=") - 1
					p_QSName = Mid(p_QSItem,1,p_QSEqual)					
						y = 0
						Do while y <= p_excludeCount
							p_excludeItem = p_excludeArray(y)							
							If lCase(p_QSName) = lCase(p_excludeItem) Then													
								p_QS = Replace(p_QS,p_QSItem,"")
								p_QS = Replace(p_QS,p_delimiter & p_delimiter,p_delimiter)											
							End If
							
							y = y + 1
						Loop
					x = x + 1
				Loop				
				p_QS = Replace(p_QS,p_delimiter,"&")
				If mid(p_QS,1,1) <> "&" Then
					p_QS = ("&" & p_QS)					
				End If
				If p_QS = "&" Then
					p_QS = ""				
	 			End If											
			End If	
			
			' Build Tabs
			p_strOutput = "<div><ul class='tabUL'>"
			
			p_tabArray = split(p_tabTitles,p_delimiter)	
			p_tabCount = uBound(p_tabArray)
			
			' Loop through tabs and build <li> tags
			x = 0
			Do while x <= p_tabCount
				p_tabTitlePre = p_tabArray(x)			
				p_tabTitlePost = Replace(p_tabArray(x),"_"," ")
				
				If Ucase(p_tabSelected) = UCase(p_tabTitlePre) Then
					p_tabState = "Selected"
				Else
					p_tabState = ""		
				End If		
				
				If p_tabType = "round" Or p_tabType = "" Then
					p_strOutput = p_strOutput & "<li class='tabLI'><a class='tabMenu' href='" & p_URL & "?selectedContent=" & p_tabTitlePre & "&amp;XMLFile=" & p_tabTitlePre & p_QS & "'><div class='tabSnazzy'><div class='tabTop'><div class='tabRnd1" & p_tabState & "'></div><div class='tabRnd2" & p_tabState & "'></div><div class='tabRnd3" & p_tabState & "'></div><div class='tabRnd4" & p_tabState & "'></div></div><span class='tabContent" & p_tabState & "'>" & p_tabTitlePost & "</span></div></a></li>"
				Else
					p_strOutput = p_strOutput & "<li class='tabLI'><a class='tabMenu' href='" & p_URL & "?selectedContent=" & p_tabTitlePre & "&amp;XMLFile=" & p_tabTitlePre & p_QS & "'><div class='tabSnazzy'><div class='tabTop'><div class='tabSqr1" & p_tabState & "'></div><div class='tabSqr2" & p_tabState & "'></div><div class='tabSqr3" & p_tabState & "'></div><div class='tabSqr4" & p_tabState & "'></div></div><span class='tabContent" & p_tabState & "'>" & p_tabTitlePost & "</span></div></a></li>"
				End If
				x = x + 1
			Loop
			
			p_strOutput = p_strOutput & "</ul></div><br /><div class='tabBtmLine'></div>"	
			
			getTabs = p_strOutput
		End Function
		
End Class
%>