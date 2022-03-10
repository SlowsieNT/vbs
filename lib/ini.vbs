' License: 0-BSD, or Unlicense
' No copyright claims
Class IniFile
	Private m_Debug
	Private m_IniBuffer, m_IniLines, m_ReadFilesize, m_Filename
	Public Sub Open(aFilename)
		m_Filename = aFilename
		m_IniBuffer = ReadFileUTF8(aFilename)
		m_ReadFilesize = Len(m_IniBuffer)
		HandleIniRead
	End Sub
	' 
	Public Function GetValue(aSection, aKey)
		'0=m_Section, 1=m_Key, 2=m_Value, 3=m_Line, 4=m_Comment
		For vI = 0 To UBound(m_IniLines)
			Dim vItem : vItem = m_IniLines(vI)
			' find key name by section name
			If LCase(vItem(1)) = LCase(aKey) And LCase(vItem(0)) = LCase(aSection) Then
				GetValue = vItem(2)
				Exit For
			End If
		Next
	End Function

	Public Function StrArrayIndexOf(aArray, aValue, aUseLCase)
		StrArrayIndexOf = -1
		For vI = 0 To UBound(aArray)
			Dim vItem : vItem = aArray(vI)
			If aUseLCase Then
				If LCase(vItem) = LCase(aValue) Then
					StrArrayIndexOf = vI
					Exit For
				End If
			Else
				If (vItem) = (aValue) Then
					StrArrayIndexOf = vI
					Exit For
				End If
			End If
		Next
	End Function

	Public Function GetSections()
		Dim vSections : vSections = Array()
		For vI = 0 To UBound(m_IniLines)
			Dim vItem : vItem = m_IniLines(vI)
			Dim vItemIdx : vItemIdx = StrArrayIndexOf(vSections, vItem(0), 0)
			If -1 = vItemIdx Then PreserveAdd vSections, vItem(0)
		Next
		GetSections = vSections
	End Function

	Public Function SetValue(aSection, aKey, aValue)
		'0=m_Section, 1=m_Key, 2=m_Value, 3=m_Line, 4=m_Comment
		Dim vFound : vFound=0
		For vI = 0 To UBound(m_IniLines)
			Dim vItem : vItem = m_IniLines(vI)
			' find key name by section name
			If LCase(vItem(1)) = LCase(aKey) And LCase(vItem(0)) = LCase(aSection) Then
				m_IniLines(vI) = Array(vItem(0),vItem(1),aValue,vItem(3),vItem(4))
				vFound = 1
				Exit For
			End If
		Next
		If 0 = vFound Then
			PreserveAdd m_IniLines, array(aSection,aKey,aValue,"","")
		End If
	End Function

	' function that should read ini:
	Sub HandleIniRead()
		Dim vLines : vLines = Split(m_IniBuffer, vbNewLine)
		Dim vSection : vSection = ""
		Dim vItemCount : vItemCount = UBound(vLines)
		Dim vItems : vItems = Array()
		Dim vSQ, vDQ: vDQ = Chr(34) : vSQ = "'"
		For vI = 0 To UBound(vLines)
			Dim vLine : vLine = vLines(vI)
			Dim vIniLine : vIniLine = Array("","","","","",0)
			Dim vCanAdd : vCanAdd = 1
			'0=m_Section, 1=m_Key, 2=m_Value, 3=m_Line, 4=m_Comment
			vIniLine(3) = vLine
			Dim vBracket : vBracket = IndexOf(vLine, "[", 0, 0)
			Dim vBracket2 : vBracket2 = IndexOf(vLine, "]", 0, 0)
			If -1 <> vBracket And -1 <> vBracket2 Then
				vSection = Mid(vLine, 2, -2+vBracket2)
				vCanAdd = 0
				PreserveAdd vItems, Array(vSection,"","",Mid(vLine, 1+vBracket2),vIniLine(4),1)
			Else
				' If not section name, parse quoted string
				Dim vInQuote: vInQuote = 0
				Dim vCommentPos: vCommentPos = -1
				Dim vBufferQ: vBufferQ = ""
				' vSQ, vDQ
				For vI2 = 1 To Len(vLine)
					Dim vChar: vChar = Mid(vLine, vI2, 1)
					Dim vIQ: vIQ = 0
					' Detect both quote types
					If vChar = vSQ Then vIQ = 1
					If vChar = vDQ Then vIQ = 2
					' vInQuote
					If 0 <> vInQuote Then
						' if same then ignore char
						If vIQ <> vInQuote Then
							vBufferQ = vBufferQ & vChar
						Else
							vInQuote = 0
						End If
					Else
						vInQuote = vIQ
						If ";" = vChar Then
							vCommentPos = vI2
							If 0 = Len(vIniLine(1)) Then vIniLine(1) = vBufferQ
							Exit For
						ElseIf "=" = vChar Then
							vIniLine(1) = vBufferQ
							If 0 = Len(vIniLine(1)) Then vIniLine(1)=null
							vBufferQ = ""
						ElseIf 0 = vIQ Then
							vBufferQ = vBufferQ & vChar
						End If
					End If
				Next
				' Handle regular string
				'0=m_Section, 1=m_Key, 2=m_Value, 3=m_Line, 4=m_Comment
				If -1 <> vCommentPos Then vIniLine(4) = Mid(vLine, 1 + vCommentPos)
				' Assign value
				vIniLine(2) = vBufferQ
				' Assign section
				vIniLine(0) = vSection
			End If
			If vCanAdd Then PreserveAdd vItems, vIniLine
		Next
		m_IniLines = vItems
	End Sub

	Function GetSectionDefinition(aSection)
		For vI = 0 To UBound(m_IniLines)
			Dim vItem : vItem = m_IniLines(vI)
			If LCase(vItem(0)) = LCase(aSection) Then
				GetSectionDefinition = vItem
				Exit For
			End If
		Next
	End Function

	Public Function GetLinesBySection(aSection)
		Dim vArr : vArr = Array()
		For vI = 0 To UBound(m_IniLines)
			Dim vItem : vItem = m_IniLines(vI)
			If LCase(vItem(0)) = LCase(aSection) Then
				PreserveAdd vArr, vItem
			End If
		Next
		GetLinesBySection = vArr
	End Function
	' aFilename optional, use 0, "", False, or Nothing as argument to default to m_Filename
	Public Sub Write(aFilename)
		Dim vOutput : vOutput = Array()
		Dim vSections : vSections = GetSections()
		Dim vFilename : vFilename = aFilename
		' didn't bother to shorten this
		If 8 <> VarType(aFilename) Or "" = aFilename Then vFilename = m_Filename
		For vI = 0 To UBound(vSections)
			Dim vSection : vSection = vSections(vI)
			Dim vSItems : vSItems = GetLinesBySection(vSection)
			If "" <> vSection Then
				Dim vSectionDef : vSectionDef = GetSectionDefinition(vSection)
				Dim vTemp : vTemp = "[" & vSection & "]"
				If "" <> vSectionDef(3) Then vTemp = vTemp & vSectionDef(3)
				If "" <> vSectionDef(4) Then vTemp = vTemp & ";" & vSectionDef(4)
				PreserveAdd vOutput, vTemp
			End If
			For vI2 = 0 To UBound(vSItems)
				Dim vILine : vILine = vSItems(vI2)
				Dim vStr: vStr = ""
				' vILine (0=m_Section, 1=m_Key, 2=m_Value, 3=m_Line, 4=m_Comment)
				If "" <> vILine(1) Then
					vStr = vILine(1) & "=" & WrapValue(vILine(2))
					If "" <> vILine(4) Then vStr = vStr & ";" & vILine(4)
					PreserveAdd vOutput, vStr
				Else
					If "" <> Trim(vILine(2)) Then vStr = "=" & WrapValue(vILine(2))
					If "" <> vILine(4) Then vStr = vStr & ";" & vILine(4)
					If "" <> vStr Then PreserveAdd vOutput, vStr
				End If
			Next
		Next
		If m_Debug Then
			MsgBox join(vOutput, vbNewLine)
		Else
			WriteFileUTF8 vFilename, join(vOutput, vbNewLine)
		End If
	End Sub

	' Standalone functions:
	Public Function WrapValue(aString)
		Dim vRet: vRet = ""
		Dim vSQ, vDQ: vDQ = Chr(34) : vSQ = "'"
		For vI2 = 1 To Len(aString)
			Dim vChar: vChar = Mid(aString, vI2, 1)
			If "=" = vChar Then vChar = "'='"
			If ";" = vChar Then vChar = "';'"
			If vDQ = vChar Then vChar = "'" & vDQ & "'"
			If vSQ = vChar Then vChar = vDQ & vSQ & vDQ
			vRet = vRet & vChar
		Next
		WrapValue = vRet
	End Function
	Sub PreserveAdd(ByRef aArray, ByVal aValue)
		Dim vUB
		If IsArray(aArray) Then
			On Error Resume Next
			vUB = UBound(aArray)
			If Err.Number <> 0 Then vUB = -1
			ReDim Preserve aArray(1 + vUB)
			aArray(UBound(aArray)) = aValue
		End If
	End Sub
	Public Function IndexOf(aHaystack, aNeedle, aUseLCase, aStartIndex)
		Dim vNeedleLen: vNeedleLen = Len(aNeedle)
		Dim vLCasedN, vIBLen: vIBLen = Len(aHaystack)
		If aUseLCase Then vLCasedN = LCase(aNeedle)
		If aStartIndex < 1 Then aStartIndex = 1
		IndexOf = -1
		For vI = aStartIndex To vIBLen
			Dim vPiece: vPiece = Mid(aHaystack, vI, vNeedleLen)
			If aUseLCase Then
				If LCase(vPiece) = vLCasedN Then
					IndexOf = vI
					Exit For
				End If
			Else
				If vPiece = aNeedle Then
					IndexOf = vI
					Exit For
				End If
			End If
		Next
	End function
	Function ReadFileUTF8(aFilename)
		On Error Resume Next
		Set vStream = CreateObject("ADODB.Stream")
		vStream.CharSet = "utf-8" : vStream.Open
		vStream.LoadFromFile(aFilename)
		ReadFileUTF8 = vStream.ReadText()
	End Function
	Function WriteFileUTF8(aFilename, aData)
		On Error Resume Next
		Set vStream = CreateObject("ADODB.Stream")
		vStream.CharSet = "utf-8" : vStream.Open
		vStream.WriteText aData : vStream.SaveToFile aFilename, 2
	End Function

End Class
