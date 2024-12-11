Option Strict Off
Option Explicit On
Module ManageINIFile
	' Copyright Rob Latour, 1998
	' February 21st, 1998
	
	Function fsINIParse(ByRef sIdentifier As String, ByRef sINILine As String) As String
		
		Dim iINILineLength As Short
		Dim iIdentiferLength As Short
		
		iINILineLength = Len(sINILine)
		iIdentiferLength = Len(sIdentifier)
		
		fsINIParse = Trim(Right(sINILine, iINILineLength - iIdentiferLength))
		
	End Function
	
	Sub ManageDefaults(ByRef sFunction As String)
		
		Const NumberOfLinesInTheINIFile As Short = 23
		Dim INIFileName As String
		Dim sINI(NumberOfLinesInTheINIFile) As String
		Dim sINILine(NumberOfLinesInTheINIFile) As String
		Dim iLineNumber As Short
		
		INIFileName = VB6.GetPath & "\Concentration.ini"
		
		sINI(1) = "[Concentration]"
		sINI(2) = " "
		sINI(3) = "Player 1 = "
		sINI(4) = "Player 1 Person Name = "
		sINI(5) = "Player 1 Computer Skill = "
		sINI(6) = " "
		sINI(7) = "Player 2 = "
		sINI(8) = "Player 2 Person Name = "
		sINI(9) = "Player 2 Computer Skill = "
		sINI(10) = " "
		sINI(11) = "Match Flip Time = "
		sINI(12) = "No Match Flip Time = "
		sINI(13) = "Time Between Computer Selections = "
		sINI(14) = "Time Between Computer Matches = "
		sINI(15) = " "
		sINI(16) = "Card Back = "
		sINI(17) = " "
		sINI(18) = "Save Options on Exit = "
		sINI(19) = " "
		sINI(20) = "Height = "
		sINI(21) = "Width  = "
		sINI(22) = "Top    = "
		sINI(23) = "Left   = "
		
		
		On Error GoTo CreateDefault
		
		Select Case sFunction
			
			Case Is = "Read"
				
				FileOpen(1, INIFileName, OpenMode.Input)
				For iLineNumber = 1 To NumberOfLinesInTheINIFile
					Input(1, sINILine(iLineNumber))
				Next iLineNumber
				FileClose(1)
				
				If fsINIParse(sINI(3), sINILine(3)) = "Person" Then
					frmOptions.DefInstance.PlayerPerson(1).Checked = True
					frmOptions.DefInstance.PlayerComputer(1).Checked = False
					
				Else
					frmOptions.DefInstance.PlayerPerson(1).Checked = False
					frmOptions.DefInstance.PlayerComputer(1).Checked = True
				End If
				
				frmOptions.DefInstance.PlayerName(1).Text = fsINIParse(sINI(4), sINILine(4))
				frmOptions.DefInstance.Slider(1).Value = CInt(fsINIParse(sINI(5), sINILine(5)))
				
				If fsINIParse(sINI(7), sINILine(7)) = "Person" Then
					frmOptions.DefInstance.PlayerPerson(2).Checked = True
					frmOptions.DefInstance.PlayerComputer(2).Checked = False
				Else
					frmOptions.DefInstance.PlayerPerson(2).Checked = False
					frmOptions.DefInstance.PlayerComputer(2).Checked = True
				End If
				
				frmOptions.DefInstance.PlayerName(2).Text = fsINIParse(sINI(8), sINILine(8))
				frmOptions.DefInstance.Slider(2).Value = CInt(fsINIParse(sINI(9), sINILine(9)))
				
				frmOptions.DefInstance.FlipTimeOnMatch.Text = fsINIParse(sINI(11), sINILine(11))
				frmOptions.DefInstance.FlipTimeOnNoMatch.Text = fsINIParse(sINI(12), sINILine(12))
				frmOptions.DefInstance.FlipTimeBetweenComputerSelections.Text = fsINIParse(sINI(13), sINILine(13))
				frmOptions.DefInstance.TimeBetweenComputerMatches.Text = fsINIParse(sINI(14), sINILine(14))
				
				frmOptions.DefInstance.lblBackOfCardName.Text = fsINIParse(sINI(16), sINILine(16))
				
				If fsINIParse(sINI(18), sINILine(18)) = "Yes" Then
					frmOptions.DefInstance.SaveOptionsOnExitCheck.CheckState = System.Windows.Forms.CheckState.Checked
				Else
					frmOptions.DefInstance.SaveOptionsOnExitCheck.CheckState = System.Windows.Forms.CheckState.Unchecked
				End If
				
				PlayTable.DefInstance.Height = VB6.TwipsToPixelsY(CSng(fsINIParse(sINI(20), sINILine(20))))
				PlayTable.DefInstance.Width = VB6.TwipsToPixelsX(CSng(fsINIParse(sINI(21), sINILine(21))))
				PlayTable.DefInstance.Top = VB6.TwipsToPixelsY(CSng(fsINIParse(sINI(22), sINILine(22))))
				PlayTable.DefInstance.Left = VB6.TwipsToPixelsX(CSng(fsINIParse(sINI(23), sINILine(23))))
				
			Case Is = "Write Everything"
				
				sINILine(1) = sINI(1)
				sINILine(2) = sINI(2)
				If frmOptions.DefInstance.PlayerPerson(1).Checked = True Then
					sINILine(3) = sINI(3) & "Person"
				Else
					sINILine(3) = sINI(3) & "Computer"
				End If
				sINILine(4) = sINI(4) & frmOptions.DefInstance.PlayerName(1).Text
				sINILine(5) = sINI(5) & Trim(Str(frmOptions.DefInstance.Slider(1).Value))
				sINILine(6) = sINI(6)
				If frmOptions.DefInstance.PlayerPerson(2).Checked = True Then
					sINILine(7) = sINI(7) & "Person"
				Else
					sINILine(7) = sINI(7) & "Computer"
				End If
				sINILine(8) = sINI(8) & frmOptions.DefInstance.PlayerName(2).Text
				sINILine(9) = sINI(9) & Trim(Str(frmOptions.DefInstance.Slider(2).Value))
				sINILine(10) = sINI(10)
				sINILine(11) = sINI(11) & frmOptions.DefInstance.FlipTimeOnMatch.Text
				sINILine(12) = sINI(12) & frmOptions.DefInstance.FlipTimeOnNoMatch.Text
				sINILine(13) = sINI(13) & frmOptions.DefInstance.FlipTimeBetweenComputerSelections.Text
				sINILine(14) = sINI(14) & frmOptions.DefInstance.TimeBetweenComputerMatches.Text
				sINILine(15) = sINI(15)
				sINILine(16) = sINI(16) & frmOptions.DefInstance.lblBackOfCardName.Text
				sINILine(17) = sINI(17)
				If frmOptions.DefInstance.SaveOptionsOnExitCheck.CheckState = 1 Then
					sINILine(18) = sINI(18) & "Yes"
				Else
					sINILine(18) = sINI(18) & "No"
				End If
				sINILine(19) = sINI(19)
				sINILine(20) = sINI(20) & Trim(Str(VB6.PixelsToTwipsY(PlayTable.DefInstance.Height)))
				sINILine(21) = sINI(21) & Trim(Str(VB6.PixelsToTwipsX(PlayTable.DefInstance.Width)))
				sINILine(22) = sINI(22) & Trim(Str(VB6.PixelsToTwipsY(PlayTable.DefInstance.Top)))
				sINILine(23) = sINI(23) & Trim(Str(VB6.PixelsToTwipsX(PlayTable.DefInstance.Left)))
				
				FileOpen(1, INIFileName, OpenMode.Output)
				For iLineNumber = 1 To NumberOfLinesInTheINIFile
					PrintLine(1, sINILine(iLineNumber))
				Next iLineNumber
				FileClose(1)
				
			Case Is = "Write Position Only"
				
				FileOpen(1, INIFileName, OpenMode.Input)
				For iLineNumber = 1 To NumberOfLinesInTheINIFile
					Input(1, sINILine(iLineNumber))
				Next iLineNumber
				FileClose(1)
				
				sINILine(20) = sINI(20) & Trim(Str(VB6.PixelsToTwipsY(PlayTable.DefInstance.Height)))
				sINILine(21) = sINI(21) & Trim(Str(VB6.PixelsToTwipsX(PlayTable.DefInstance.Width)))
				sINILine(22) = sINI(22) & Trim(Str(VB6.PixelsToTwipsY(PlayTable.DefInstance.Top)))
				sINILine(23) = sINI(23) & Trim(Str(VB6.PixelsToTwipsX(PlayTable.DefInstance.Left)))
				
				FileOpen(1, INIFileName, OpenMode.Output)
				For iLineNumber = 1 To NumberOfLinesInTheINIFile
					PrintLine(1, sINILine(iLineNumber))
				Next iLineNumber
				FileClose(1)
				
		End Select
		
		GoTo EndLoadINIFIle
		
CreateDefault: 
		
		On Error GoTo 0
		ManageDefaults(("Write Everything"))
		
EndLoadINIFIle: 
		
	End Sub
End Module