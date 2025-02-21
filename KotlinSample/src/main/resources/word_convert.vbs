
' -- Functions -----------------------------------------------------------------
Function ParseJSONArray(jsonText)
	Dim dict, itemDict, itemKey, jsonItems, jsonItem, i

	jsonText = Replace(jsonText, "[", "")
	jsonText = Replace(jsonText, "]", "")
	jsonItems = Split(jsonText, "},{")

	Set dict = CreateObject("Scripting.Dictionary")

	For i = 0 To UBound(jsonItems)
		jsonItem = "{" & Replace(jsonItems(i), "{", "") & "}"
		Set itemDict = ParseJSON(jsonItem)
		itemKey = itemDict("key")
		If itemKey <> "" AND dict.Exists(itemKey) = False Then
			dict.Add itemKey, itemDict
		End If
	Next
	Set ParseJSONArray = dict

End Function

Function ParseJSON(jsonText)
	Dim dict, pairs, pair, key, value, i

	jsonText = Replace(jsonText, "{", "")
	jsonText = Replace(jsonText, "}", "")
	pairs = Split(jsonText, ",")

	Set dict = CreateObject("Scripting.Dictionary")

	Dim emptyKey
	emptyKey = False
	For i = 0 To UBound(pairs)
		If emptyKey = True Then
			emptyKey = False
		Else
			pair = Split(pairs(i), ":")
			key = Trim(Replace(pair(0), """", ""))
			value = Trim(Replace(pair(1), """", ""))
			If key = "key" AND value = "" Then
				emptyKey = True
			Else
				dict.Add key, value
			End If
		End If
    Next

    Set ParseJSON = dict
End Function

Function RemoveWhitespace(text)
	text = Replace(text, vbCrLf, "")
	text = Replace(text, vbCr, "")
	text = Replace(text, vbLf, "")
	text = Replace(text, vbTab, "")

	RemoveWhitespace = text
End Function

Function SortDictionaryKeys(d)
	Dim arrKeys, i, j, temp
	arrKeys = d.Keys

	For i = 0 To UBound(arrKeys) - 1
		For j = i + 1 To UBound(arrKeys)
			If (Len(arrKeys(i)) < Len(arrKeys(j))) Or _
			   ((Len(arrKeys(i)) = Len(arrKeys(j))) And (arrKeys(i) > arrKeys(j))) Then
				temp = arrKeys(i)
				arrKeys(i) = arrKeys(j)
				arrKeys(j) = temp
			End If
		Next
	Next

	SortDictionaryKeys = arrKeys
End Function

' ---------------------------------------------------------------------------------


Dim fileSystemObject
Dim inputFile, outputFile, outputFormat

inputFile = WScript.Arguments.Unnamed.Item(0)
outputFile = WScript.Arguments.Unnamed.Item(1)
outputFormat = CInt(WScript.Arguments.Unnamed.Item(2)) 'word = 16

Set fileSystemObject = CreateObject("Scripting.FileSystemObject")
inputFile = fileSystemObject.GetAbsolutePathName(inputFile)

If fileSystemObject.FileExists(inputFile) Then

	On Error Resume Next
	
	Set WA = CreateObject("Word.Application")
	Set WD = WA.Documents.Open(inputFile)
	If WD = "" OR Err.Number <> 0 Then
		'WScript.Echo "Open Failed : " & inputFile & vbCrLf & Err.Description
		'WA.Quit
		WScript.Quit -2
	End If
	On Error GoTo 0

	On Error Resume Next
	Set ActiveDocument = WA.ActiveDocument
	Set Selection = WA.Selection
	
	Dim jsonData, cleanJson, sortedKeys
	jsonData = "@{#JSONDATA#}@"
	
	On Error Resume Next
	cleanJson = RemoveWhitespace(jsonData)
	Set jsonDict = ParseJSONArray(cleanJson)
	sortedKeys = SortDictionaryKeys(jsonDict)

	If Err.Number <> 0 Then
		WD.Close 0
		WA.Quit
		'WScript.Quit -2
	End If
	On Error GoTo 0

	On Error Resume Next
	Selection.Find.ClearFormatting

	Dim shapeMargin
	shapeMargin = 2

	' Search FormID..
	Dim itemDict, formID, compType, bError
	For Each searchKey In sortedKeys
		Selection.Find.ClearFormatting
		Selection.Range.Start = 0
		Selection.Range.End = 0
		Selection.Collapse 1
		
		'WScript.Echo "searchKey : " & searchKey

		Set itemDict = jsonDict(searchKey)
		formID = itemDict("formid")
		compType = "label"
		For Each story In ActiveDocument.StoryRanges
			Err.Clear
			If story.StoryType = 1 Or story.StoryType = 5 Or story.StoryType = 7 Or story.StoryType = 9 Then
				Do
					bError = False
					Do
						story.Find.ClearFormatting
						With story.Find
							.Text = searchKey
							.Forward = True
							.Wrap = 1
							.Format = False
							.MatchCase = True
							.MatchWholeWord = True
							.MatchAllWordForms = False
							.MatchSoundsLike = False
							.MatchWildcards = False
							.MatchPunctuation = True
							.MatchByte = True
							.Execute
						End With
						
						If story.Find.Found Then
							story.Select
							
							Dim selEnd, paraEnd, paraAlign
							selEnd = Selection.Range.End
							paraEnd = Selection.Paragraphs(1).Range.End - 1
							paraAlign = Selection.ParagraphFormat.Alignment
							'WScript.Echo "selStart=" & selStart & ", paraStart=" & paraStart
							
							Dim t: t = story.Text
							If t = searchKey Then
								'WScript.Echo "convert : " & t
								On Error Resume Next
								Dim fontName, fontSize, fontBold, fontItalic, fontColor
								fontName = Selection.Range.Font.Name
								fontSize = Selection.Range.Font.Size
								fontBold = Selection.Range.Font.Bold
								fontItalic = Selection.Range.Font.Italic
								fontColor = Selection.Range.Font.Color
								
								Dim Left, Width, Top, Right, posX
								Set Starting = Selection.Range
								Starting.SetRange Starting.Start, Starting.Start
								
								'WA.ActiveWindow.ScrollIntoView Selection.Range, True
								
								Left = Starting.Information(5)' - shapeMargin
								Top = Starting.Information(6)
								
								If Left = "" or Top = "" Then
									'WScript.Echo "Error position.."
									bError = True
									Exit Do
								End If
								
								Set Ending = Selection.Range
								Ending.SetRange Ending.End, Ending.End
								Width = (Ending.Information(5) - Left)' + shapeMargin * 2
								
								If Err.Number <> 0 Then
									'WScript.Echo "Error 1 : " & Err.Description
									bError = True
									Exit Do
								End If
								On Error GoTo 0
								
								On Error Resume Next
								Dim tempShape, alterText
								If story.StoryType = 7 Then 'Header Story
									Set header = ActiveDocument.Sections(1).Headers(1) ' wdHeaderFooterPrimary = 1
									Set tempShape = header.Shapes.AddTextbox(1, Left, Top, Width, 200)
								ElseIf story.StoryType = 9 Then 'Footer Story
									Set footer = ActiveDocument.Sections(1).Footers(1) ' wdHeaderFooterPrimary = 1
									Set tempShape = footer.Shapes.AddTextbox(1, Left, Top, Width, 200)
								Else 'Main & Etc
									Set tempShape = ActiveDocument.Shapes.AddTextbox(1, Left, Top, Width, 200)
								End If
								tempShape.Name = formID
								tempShape.Fill.Visible = False
								alterText = "#OZBEGIN#" + vbCrLf + "{" + vbCrLf + """type"":""" & compType & """," + vbCrLf + """formID"":""" & formID & """" + vbCrLf + "}" + vbCrLf + "#OZEND#"
								tempShape.AlternativeText = alterText
								
								With tempShape.TextFrame
									.MarginTop = 0
									.MarginBottom = 0
									.MarginLeft = 0
									.MarginRight = 0
									.VerticalAnchor = 3
									.TextRange.Font.Name = fontName
									.TextRange.Font.Size = fontSize
									.TextRange.Font.Bold = fontBold
									.TextRange.Font.Italic = fontItalic
									.TextRange.Font.Color = fontColor
									.TextRange.ParagraphFormat.Alignment = 1
									.TextRange.Text = t
									.AutoSize = True
								End With
								
								With tempShape
									.Line.Weight = 0.5
									.Line.ForeColor.RGB = RGB(255, 127, 39)
									.Line.DashStyle = 8
								End With
								
								tempShape.TextFrame.TextRange.Text = ""	
								
								If Err.Number <> 0 Then
									'WScript.Echo "Error 2 : " & Err.Description
									bError = True
									Exit Do
								End If
								
								Selection.SetRange Starting.Start, Ending.End
								
								If paraAlign = 2 Then
									Dim spaceChar
									If selEnd = paraEnd Then
										spaceChar = ChrW(160)
									Else
										spaceChar = " "
									End If
									
									'WScript.Echo "selEnd=" & selEnd & ", paraEnd=" & paraEnd
									
									Ending.SetRange Selection.Range.Start, Selection.Range.Start
									Right = Ending.Information(5)
									Selection.Delete
									
									Do
										Set Starting2 = Selection.Range
										Starting2.SetRange Starting2.Start, Starting2.Start
										posX = Starting2.Information(5)
										'WScript.Echo "Left=" & posX & ", Right=" & Right
										
										If posX > Right Then
											Selection.InsertBefore spaceChar
											Selection.Move 1, -1
										Else
											Selection.Delete 1, 1
											Err.Clear
											Exit Do
										End If
									Loop
								Else
									Ending.SetRange Selection.Range.End, Selection.Range.End
									Right = Ending.Information(5)
									Selection.Delete
									
									Do
										Set Starting2 = Selection.Range
										Starting2.SetRange Starting2.Start, Starting2.Start
										posX = Starting2.Information(5)
										If posX < Right Then
											Selection.InsertAfter " "
											Selection.Move 1
										Else
											Selection.Move 1, -1
											Selection.Delete 1, 1
											Err.Clear
											Exit Do
										End If
									Loop
								End If
								
								If Err.Number <> 0 Then
									'WScript.Echo "Error 4 : " & Err.Description
									bError = True
									Exit Do
								End If
								
							Else
								bError = True
								Exit Do
							End If
							'Selection.Move 1
							Selection.Collapse
							Selection.SetRange ActiveDocument.Range.Start, ActiveDocument.Range.Start
						Else
							Exit Do
						End If
					Loop

				'Selection.Move 1
				Selection.Collapse
				If bError = True Then
					'WScript.Echo "error=true : " & vbCrLf & Err.Description
					Exit Do
				End If
				Err.Clear
				Set story = story.NextStoryRange
				Loop While Not story Is Nothing
			End If
		Next
		
		Set itemDict = Nothing
	Next
	
	Selection.Find.ClearFormatting
		
	If Err.Number <> 0 Then
		'WScript.Echo "Save Failed : " & outputFile & vbCrLf & Err.Description
		WScript.Quit -3
		'WA.Quit
	Else
		WD.SaveAs outputFile, outputFormat
		'WScript.Echo "Save Success! : " & outputFile
	End If
	WD.Close 0
	On Error GoTo 0
Else
	'WScript.Echo "Files does not exist : " & inputFile & vbCrLf & Err.Description
	WScript.Quit -4
	'WA.Quit
End If

WScript.Quit 2
'WA.Quit
' ---- END



