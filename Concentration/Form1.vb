Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class PlayTable
	Inherits System.Windows.Forms.Form
#Region "Windows Form Designer generated code "
	Public Sub New()
		MyBase.New()
		If m_vb6FormDefInstance Is Nothing Then
			If m_InitializingDefInstance Then
				m_vb6FormDefInstance = Me
			Else
				Try 
					'For the start-up form, the first instance created is the default instance.
					If System.Reflection.Assembly.GetExecutingAssembly.EntryPoint.DeclaringType Is Me.GetType Then
						m_vb6FormDefInstance = Me
					End If
				Catch
				End Try
			End If
		End If
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			Static fTerminateCalled As Boolean
			If Not fTerminateCalled Then
				Form_Terminate_renamed()
				fTerminateCalled = True
			End If
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents Timer2 As System.Windows.Forms.Timer
	Public WithEvents CommonDialog1 As AxMSComDlg.AxCommonDialog
	Public WithEvents Timer1 As System.Windows.Forms.Timer
	Public WithEvents StatusBar1 As AxComctlLib.AxStatusBar
	Public WithEvents WinnersBox As System.Windows.Forms.Label
	Public WithEvents MenuGameDeal As System.Windows.Forms.MenuItem
	Public WithEvents MenuGameOptions As System.Windows.Forms.MenuItem
	Public WithEvents MenuGameExit As System.Windows.Forms.MenuItem
	Public WithEvents MenuGame As System.Windows.Forms.MenuItem
	Public WithEvents MenuHelpHelpTopics As System.Windows.Forms.MenuItem
	Public WithEvents MenuHelpAbout As System.Windows.Forms.MenuItem
	Public WithEvents MenuHelp As System.Windows.Forms.MenuItem
	Public MainMenu1 As System.Windows.Forms.MainMenu
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(PlayTable))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ToolTip1.Active = True
		Me.Timer2 = New System.Windows.Forms.Timer(components)
		Me.CommonDialog1 = New AxMSComDlg.AxCommonDialog
		Me.Timer1 = New System.Windows.Forms.Timer(components)
		Me.StatusBar1 = New AxComctlLib.AxStatusBar
		Me.WinnersBox = New System.Windows.Forms.Label
		Me.MainMenu1 = New System.Windows.Forms.MainMenu
		Me.MenuGame = New System.Windows.Forms.MenuItem
		Me.MenuGameDeal = New System.Windows.Forms.MenuItem
		Me.MenuGameOptions = New System.Windows.Forms.MenuItem
		Me.MenuGameExit = New System.Windows.Forms.MenuItem
		Me.MenuHelp = New System.Windows.Forms.MenuItem
		Me.MenuHelpHelpTopics = New System.Windows.Forms.MenuItem
		Me.MenuHelpAbout = New System.Windows.Forms.MenuItem
		CType(Me.CommonDialog1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.StatusBar1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
		Me.BackColor = System.Drawing.Color.FromARGB(0, 128, 0)
		Me.Text = "Concentration"
		Me.ClientSize = New System.Drawing.Size(621, 405)
		Me.Location = New System.Drawing.Point(7, 67)
		Me.Icon = CType(resources.GetObject("PlayTable.Icon"), System.Drawing.Icon)
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MaximizeBox = True
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "PlayTable"
		Me.Timer2.Enabled = False
		Me.Timer2.Interval = 100
		CommonDialog1.OcxState = CType(resources.GetObject("CommonDialog1.OcxState"), System.Windows.Forms.AxHost.State)
		Me.CommonDialog1.Location = New System.Drawing.Point(50, 272)
		Me.CommonDialog1.Name = "CommonDialog1"
		Me.Timer1.Enabled = False
		Me.Timer1.Interval = 500
		StatusBar1.OcxState = CType(resources.GetObject("StatusBar1.OcxState"), System.Windows.Forms.AxHost.State)
		Me.StatusBar1.Dock = System.Windows.Forms.DockStyle.Bottom
		Me.StatusBar1.Size = New System.Drawing.Size(621, 20)
		Me.StatusBar1.Location = New System.Drawing.Point(0, 385)
		Me.StatusBar1.TabIndex = 0
		Me.StatusBar1.Name = "StatusBar1"
		Me.WinnersBox.BackColor = System.Drawing.Color.FromARGB(0, 128, 0)
		Me.WinnersBox.Text = " "
		Me.WinnersBox.Font = New System.Drawing.Font("Arial", 24!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.WinnersBox.Size = New System.Drawing.Size(631, 153)
		Me.WinnersBox.Location = New System.Drawing.Point(32, 22)
		Me.WinnersBox.TabIndex = 1
		Me.WinnersBox.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.WinnersBox.Enabled = True
		Me.WinnersBox.ForeColor = System.Drawing.SystemColors.ControlText
		Me.WinnersBox.Cursor = System.Windows.Forms.Cursors.Default
		Me.WinnersBox.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.WinnersBox.UseMnemonic = True
		Me.WinnersBox.Visible = True
		Me.WinnersBox.AutoSize = False
		Me.WinnersBox.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.WinnersBox.Name = "WinnersBox"
		Me.MenuGame.Text = "&Game"
		Me.MenuGame.Checked = False
		Me.MenuGame.Enabled = True
		Me.MenuGame.Visible = True
		Me.MenuGame.MDIList = False
		Me.MenuGameDeal.Text = "&Deal"
		Me.MenuGameDeal.Shortcut = System.Windows.Forms.Shortcut.F2
		Me.MenuGameDeal.Checked = False
		Me.MenuGameDeal.Enabled = True
		Me.MenuGameDeal.Visible = True
		Me.MenuGameDeal.MDIList = False
		Me.MenuGameOptions.Text = "&Options"
		Me.MenuGameOptions.Shortcut = System.Windows.Forms.Shortcut.F4
		Me.MenuGameOptions.Checked = False
		Me.MenuGameOptions.Enabled = True
		Me.MenuGameOptions.Visible = True
		Me.MenuGameOptions.MDIList = False
		Me.MenuGameExit.Text = "&Exit"
		Me.MenuGameExit.Shortcut = System.Windows.Forms.Shortcut.F3
		Me.MenuGameExit.Checked = False
		Me.MenuGameExit.Enabled = True
		Me.MenuGameExit.Visible = True
		Me.MenuGameExit.MDIList = False
		Me.MenuHelp.Text = "&Help"
		Me.MenuHelp.Checked = False
		Me.MenuHelp.Enabled = True
		Me.MenuHelp.Visible = True
		Me.MenuHelp.MDIList = False
		Me.MenuHelpHelpTopics.Text = "Help &Topics"
		Me.MenuHelpHelpTopics.Shortcut = System.Windows.Forms.Shortcut.F1
		Me.MenuHelpHelpTopics.Checked = False
		Me.MenuHelpHelpTopics.Enabled = True
		Me.MenuHelpHelpTopics.Visible = True
		Me.MenuHelpHelpTopics.MDIList = False
		Me.MenuHelpAbout.Text = "&About Concentration"
		Me.MenuHelpAbout.Checked = False
		Me.MenuHelpAbout.Enabled = True
		Me.MenuHelpAbout.Visible = True
		Me.MenuHelpAbout.MDIList = False
		Me.Controls.Add(CommonDialog1)
		Me.Controls.Add(StatusBar1)
		Me.Controls.Add(WinnersBox)
		CType(Me.StatusBar1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.CommonDialog1, System.ComponentModel.ISupportInitialize).EndInit()
		Me.MenuGame.Index = 0
		Me.MenuHelp.Index = 1
		MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem(){Me.MenuGame, Me.MenuHelp})
		Me.MenuGameDeal.Index = 0
		Me.MenuGameOptions.Index = 1
		Me.MenuGameExit.Index = 2
		MenuGame.MenuItems.AddRange(New System.Windows.Forms.MenuItem(){Me.MenuGameDeal, Me.MenuGameOptions, Me.MenuGameExit})
		Me.MenuHelpHelpTopics.Index = 0
		Me.MenuHelpAbout.Index = 1
		MenuHelp.MenuItems.AddRange(New System.Windows.Forms.MenuItem(){Me.MenuHelpHelpTopics, Me.MenuHelpAbout})
		Me.Menu = MainMenu1
	End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As PlayTable
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As PlayTable
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New PlayTable()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 
	' Copyright Rob Latour 1998 - 2002
	' February 21st, 1998
	' July 28th, 2001
	
	Const CardInPlay As Short = 0
	Const CardOutOfPlay As Short = 1
	Const CardHasBeenSeen As Short = 0
	Const CardHasNotBeenSeen As Short = 1
	
	Dim Deck(52, 8) As Integer
	' (x, 0) ; card's unique value 1..52
	' (x, 1) ; card's min x position
	' (x, 2) ; card's max x position
	' (x, 3) ; card's min y position
	' (x, 4) ; card's max y postion
	' (x, 5) ; card's back or face
	' (x, 6) ; card's nominal value
	' (x, 7) ; card is 'In Play' or 'Out of Play'
	' (x, 8) ; card has been flipped or not
	
	Dim HoldOptions(4) As Short
	
	Dim Player As Short
	Dim PlayerScore(2) As Short
	Dim UpdateSpace As Short
	Dim CardsMatch As Boolean
	Dim GameInProgress As Boolean
	Dim NewGameLockout As Boolean
	Dim NewGameRequested As Boolean
	Private Sub Delay(ByVal SecondsToWait As Double)
		
		Dim dblEndTime As Double
		
		dblEndTime = VB.Timer() + SecondsToWait
		
		' Do nothing but allow other applications to process their events.
		Do While dblEndTime > VB.Timer()
			System.Windows.Forms.Application.DoEvents()
		Loop 
		
	End Sub
	Sub ShuffleTheDeck()
		
		Dim Card As Short
		Dim X As Short
		
		' shuffle the deck
		Randomize()
		For Card = 1 To 52
			Deck(Card, 0) = -1
			While Deck(Card, 0) < 0
				Deck(Card, 0) = Int(52 * Rnd())
				For X = 1 To Card - 1
					If Deck(X, 0) = Deck(Card, 0) Then
						Deck(Card, 0) = -1
					End If
				Next X
			End While
		Next Card
		
		' assign nominal values
		For Card = 1 To 52
			Deck(Card, 6) = Int(Deck(Card, 0) / 4)
		Next Card
		
	End Sub
	Private Sub SizePlayTable()
		
		Const NumberOfRows As Short = 6
		Const NumberOfColumns As Short = 10
		
		Dim HeightOption1, WidthOption1 As Object
		Dim AreaOption1 As Short
		Dim HeightOption2, WidthOption2 As Object
		Dim AreaOption2 As Short
		Dim Card As Short
		Dim Column As Object
		Dim Row As Short
		Dim ColumnOffset As Object
		Dim RowOffset As Short
		
		Dim Ret As Integer
		
		' option 1: size to height
		'UPGRADE_WARNING: Couldn't resolve default property of object HeightOption1. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		HeightOption1 = Int((ClientRectangle.Height - StatusBar1.Height) / NumberOfRows)
		'UPGRADE_WARNING: Couldn't resolve default property of object HeightOption1. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object WidthOption1. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		WidthOption1 = HeightOption1 / OriginalRatio
		'UPGRADE_WARNING: Couldn't resolve default property of object WidthOption1. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object HeightOption1. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		AreaOption1 = HeightOption1 * WidthOption1
		
		' option 2: size to width
		'UPGRADE_WARNING: Couldn't resolve default property of object WidthOption2. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		WidthOption2 = Int(ClientRectangle.Width / (NumberOfColumns))
		'UPGRADE_WARNING: Couldn't resolve default property of object WidthOption2. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object HeightOption2. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		HeightOption2 = WidthOption2 * OriginalRatio
		'UPGRADE_WARNING: Couldn't resolve default property of object WidthOption2. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: Couldn't resolve default property of object HeightOption2. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		AreaOption2 = HeightOption2 * WidthOption2
		
		' size to the optin which provides the smallest card area
		If AreaOption1 < AreaOption2 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object HeightOption1. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			nHeight = HeightOption1
			'UPGRADE_WARNING: Couldn't resolve default property of object WidthOption1. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			nWidth = WidthOption1
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object HeightOption2. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			nHeight = HeightOption2
			'UPGRADE_WARNING: Couldn't resolve default property of object WidthOption2. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			nWidth = WidthOption2
		End If
		
		' centre play area
		'UPGRADE_WARNING: Couldn't resolve default property of object ColumnOffset. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ColumnOffset = Int((ClientRectangle.Width - (nWidth * NumberOfColumns)) / 2)
		RowOffset = Int((ClientRectangle.Height - StatusBar1.Height - (nHeight * NumberOfRows)) / 2)
		
		Card = 1
		For Row = 0 To NumberOfRows - 1
			For Column = 0 To NumberOfColumns - 1
				'UPGRADE_WARNING: Couldn't resolve default property of object Column. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				If Not ((Row < 2 Or Row > 3) And (Column = 0 Or Column = 9)) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object ColumnOffset. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Column. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
					Deck(Card, 1) = Column * nWidth + ColumnOffset
					'UPGRADE_WARNING: Couldn't resolve default property of object ColumnOffset. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Column. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
					Deck(Card, 2) = (Column + 1) * nWidth + ColumnOffset - 1
					Deck(Card, 3) = Row * nHeight + RowOffset
					Deck(Card, 4) = (Row + 1) * nHeight + RowOffset - 1
					Card = Card + 1
				End If
			Next 
		Next 
		
	End Sub
	Private Sub DefinePlayTable()
		
		Const NumberOfRows As Short = 6
		Const NumberOfColumns As Short = 10
		
		Dim Card, Column As Object
		Dim Row As Short
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Card. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Card = 1
		For Row = 0 To NumberOfRows - 1
			For Column = 0 To NumberOfColumns - 1
				'UPGRADE_WARNING: Couldn't resolve default property of object Column. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				If Not ((Row < 2 Or Row > 3) And (Column = 0 Or Column = 9)) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Card. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
					Deck(Card, 5) = C_BACKS
					'UPGRADE_WARNING: Couldn't resolve default property of object Card. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
					Deck(Card, 7) = CardInPlay
					'UPGRADE_WARNING: Couldn't resolve default property of object Card. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
					Deck(Card, 8) = CardHasNotBeenSeen
					'UPGRADE_WARNING: Couldn't resolve default property of object Card. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
					Card = Card + 1
				End If
			Next 
		Next 
		
	End Sub
	Sub DrawPlayTable(ByRef sFunction As String)
		
		Dim Card As Short
        Dim Ret As Integer
        Dim hdc As Long
		
		Call SizePlayTable()
		
		If sFunction = "Clear" Then
			'UPGRADE_ISSUE: Form method PlayTable.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
            'PlayTable.Cls()
			'UPGRADE_ISSUE: Form method PlayTable.Line was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
            'Me.Line (0, 0) - (VB6.PixelsToTwipsX(PlayTable.DefInstance.Width), VB6.PixelsToTwipsY(PlayTable.DefInstance.Height)), System.Drawing.ColorTranslator.ToOle(PlayTable.DefInstance.BackColor), BF
		End If
		
		For Card = 1 To 52
			If Deck(Card, 7) = CardInPlay Then
				If Deck(Card, 5) = C_BACKS Then
					'UPGRADE_ISSUE: Form property PlayTable.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
					Ret = cdtDrawExt(hdc, Deck(Card, 1), Deck(Card, 3), nWidth, nHeight, SelectedBack, Deck(Card, 5), &HFFFF)
				Else
					'UPGRADE_ISSUE: Form property PlayTable.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
					Ret = cdtDrawExt(hdc, Deck(Card, 1), Deck(Card, 3), nWidth, nHeight, Deck(Card, 0), Deck(Card, 5), &HFFFF)
				End If
			Else
			End If
		Next 
		
	End Sub
	Sub ComputerPlayStrategyWin()
		' Stragegy:
		'
		' Look through known cards for a match and if two matching cards are found
		' then select them
		'
		' if a match is not found
		'
		' select a card at random and then look through history for its match, if a
		' a match is found then select the match
		'
		' if a match is not found
		'
		' select another random card and hope for the best
		
		Dim SeenHistory(52) As Short
		Dim NotSeenHistory(52) As Short
		Dim Seen As Object
		Dim NotSeen As Short
		Dim X As Object
		Dim Y As Short
		
		Dim strace As String
		strace = "A"
		
		ComputerCard1 = 0
		ComputerCard2 = 0
		For X = 1 To 52
			'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			For Y = (X + 1) To 52
				'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				If (Deck(X, 8) = CardHasBeenSeen) And (Deck(X, 7) = CardInPlay) Then
					If Deck(Y, 8) = CardHasBeenSeen And (Deck(Y, 7) = CardInPlay) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
						If X <> Y Then
							'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
							If Deck(X, 6) = Deck(Y, 6) Then
								'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
								ComputerCard1 = X
								ComputerCard2 = Y
								Exit For
							End If
						End If
					End If
				End If
			Next Y
		Next X
		strace = strace & "B"
		
		If ComputerCard2 > 0 Then
			GoTo AMatchWasFound
		End If
		
		' Build history of all Selection which have been seen and which have not been seen
		strace = strace & "c"
		'UPGRADE_WARNING: Couldn't resolve default property of object Seen. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Seen = 0
		NotSeen = 0
		For X = 1 To 52
			'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			If Deck(X, 7) = CardInPlay Then
				'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				If Deck(X, 8) = CardHasBeenSeen Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Seen. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
					Seen = Seen + 1
					'UPGRADE_WARNING: Couldn't resolve default property of object Seen. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
					'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
					SeenHistory(Seen) = X
				Else
					NotSeen = NotSeen + 1
					'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
					NotSeenHistory(NotSeen) = X
				End If
			End If
		Next X
		strace = strace & "d"
		'UPGRADE_WARNING: Couldn't resolve default property of object Seen. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (NotSeen = 0) And (Seen = 0) Then
			GoTo AMatchWasFound
		End If
		strace = strace & "e"
		' if there is at least one card in the not seen history
		'       pick a random selection from the not seen history
		' otherwise
		'       pick a random selection from the seen history
		
		If NotSeen > 0 Then
			ComputerCard1 = NotSeenHistory(Int(NotSeen * Rnd()) + 1)
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object Seen. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			ComputerCard1 = SeenHistory(Int(Seen * Rnd()) + 1)
		End If
		strace = strace & "f"
		' look for its match in history
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Seen. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		For X = 1 To Seen
			'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			If ComputerCard1 = SeenHistory(X) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				ComputerCard2 = X
				Exit For
			End If
		Next X
		strace = strace & "g"
		If ComputerCard2 > 0 Then
			GoTo AMatchWasFound
		End If
		strace = strace & "h"
		' Card One has already been chosen
		
		' if there is more than one card in the not seen history then
		' select another random card from the not seen history and hope for the best
		' otherwise select a card from the seen history
		
		' in either case ensure the same card as the first isn't picked
		
		If NotSeen > 1 Then
			strace = strace & "i"
			ComputerCard2 = NotSeenHistory(Int(NotSeen * Rnd()) + 1)
			While ComputerCard1 = ComputerCard2
				ComputerCard2 = NotSeenHistory(Int(NotSeen * Rnd()) + 1)
			End While
		Else
			strace = strace & "j"
			'UPGRADE_WARNING: Couldn't resolve default property of object Seen. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			ComputerCard2 = SeenHistory(Int(Seen * Rnd()) + 1)
			While ComputerCard1 = ComputerCard2
				'UPGRADE_WARNING: Couldn't resolve default property of object Seen. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				ComputerCard2 = SeenHistory(Int(Seen * Rnd()) + 1)
			End While
		End If
		
AMatchWasFound: 
		
		'Print strace
	End Sub
	Sub ComputerPlayStrategyLoose()
		' Stragegy: Pick any two of the remaining cards in the deck
		
		Dim History(52) As Short
		Dim Selection1 As Object
		Dim Selection2 As Short
		Dim X As Object
		Dim Y As Short
		
		' Build a history of all Selections still in play
		Y = 0
		For X = 1 To 52
			'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			If Deck(X, 7) = CardInPlay Then
				Y = Y + 1
				'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				History(Y) = X
			End If
		Next 
		
		' Pick the first and second Selections at random
		'UPGRADE_WARNING: Couldn't resolve default property of object Selection1. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		Selection1 = Int(Y * Rnd()) + 1
		Selection2 = Int(Y * Rnd()) + 1
		
		' Make sure the first is not the same as the second
		'UPGRADE_WARNING: Couldn't resolve default property of object Selection1. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		While Selection1 = Selection2
			Selection2 = Int(Y * Rnd()) + 1
		End While
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Selection1. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ComputerCard1 = History(Selection1)
		ComputerCard2 = History(Selection2)
		
	End Sub
	Sub ComputerPlay()
		
		Dim card1 As Object
		Dim card2 As Short
		
		NewGameLockout = True
		
		If frmOptions.DefInstance.Slider(Player)._Value > ((10 * Rnd())) Then
			Call ComputerPlayStrategyWin()
		Else
			Call ComputerPlayStrategyLoose()
		End If
		
		Call ShowCard(ComputerCard1)
		Call Delay(CDbl(frmOptions.DefInstance.FlipTimeBetweenComputerSelections.Text))
		Call ShowCard(ComputerCard2)
		Call EvalutePlayTable()
		
		NewGameLockout = False
		
		If NewGameRequested Then
			Call NewGame()
		End If
		
	End Sub
	Sub NewGame()
		
		If NewGameLockout Then
			If NewGameRequested Then
				Exit Sub
			Else
				HoldOptions(1) = CShort(frmOptions.DefInstance.FlipTimeOnMatch.Text)
				HoldOptions(2) = CShort(frmOptions.DefInstance.FlipTimeOnNoMatch.Text)
				HoldOptions(3) = CShort(frmOptions.DefInstance.FlipTimeBetweenComputerSelections.Text)
				HoldOptions(4) = CShort(frmOptions.DefInstance.TimeBetweenComputerMatches.Text)
				frmOptions.DefInstance.FlipTimeOnMatch.Text = CStr(0)
				frmOptions.DefInstance.FlipTimeOnNoMatch.Text = CStr(0)
				frmOptions.DefInstance.FlipTimeBetweenComputerSelections.Text = CStr(0)
				frmOptions.DefInstance.TimeBetweenComputerMatches.Text = CStr(0)
				NewGameRequested = True
				Exit Sub
			End If
		End If
		
		If NewGameRequested Then
			NewGameRequested = False
			frmOptions.DefInstance.FlipTimeOnMatch.Text = CStr(HoldOptions(1))
			frmOptions.DefInstance.FlipTimeOnNoMatch.Text = CStr(HoldOptions(2))
			frmOptions.DefInstance.FlipTimeBetweenComputerSelections.Text = CStr(HoldOptions(3))
			frmOptions.DefInstance.TimeBetweenComputerMatches.Text = CStr(HoldOptions(4))
		End If
		
		Player = 1
		PlayerScore(1) = 0
		PlayerScore(2) = 0
		WinnersBox.Visible = False
		GameInProgress = False
		Timer1.Enabled = False
		Timer2.Enabled = False
		
		Call ShuffleTheDeck()
		Call DefinePlayTable()
		Call DrawPlayTable("Clear")
		Call UpdateStatusBar()
		
		GameInProgress = True
		
		If frmOptions.DefInstance.PlayerComputer(Player).Checked Then
			Timer1.Enabled = True
		End If
		
	End Sub
	Sub LocateCard(ByRef X As Single, ByRef Y As Single, ByRef Card As Short, ByRef CardFound As Boolean)
		
		' (x,0) ; deck selection value 53 .. 65
		' (x,1) ; card's min x position            min x & y -> +----+
		' (x,2) ; card's max x position                         !    !
		' (x,3) ; card's min y position                         !    !
		' (x,4) ; card's max y position                         +----+ <- max x & y
		' (x,5) ; card's name
		
		CardFound = False
		
		For Card = 1 To 52
			If X >= Deck(Card, 1) Then
				If X <= Deck(Card, 2) Then
					If Y >= Deck(Card, 3) Then
						If Y <= Deck(Card, 4) Then
							If Deck(Card, 7) = CardInPlay Then
								CardFound = True
								Exit For
							End If
						End If
					End If
				End If
			End If
		Next 
		
	End Sub
	Sub ShowCard(ByVal Card As Short)
		
        Dim Ret As Integer
        Dim hdc As Long
		
		Deck(Card, 5) = C_FACES
		Deck(Card, 8) = CardHasBeenSeen
		
		'UPGRADE_ISSUE: Form property PlayTable.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		Ret = cdtDrawExt(hdc, Deck(Card, 1), Deck(Card, 3), nWidth, nHeight, Deck(Card, 0), Deck(Card, 5), &HFFFF)
		
	End Sub
	Sub HideCard(ByVal Card As Short)
		
		Dim Ret As Integer
        Dim hdc As Long
		Deck(Card, 5) = C_BACKS
		
		'UPGRADE_ISSUE: Form property PlayTable.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		Ret = cdtDrawExt(hdc, Deck(Card, 1), Deck(Card, 3), nWidth, nHeight, SelectedBack, Deck(Card, 5), &HFFFF)
		
	End Sub
	Sub RemoveCard(ByVal Card As Short)
        Dim hdc As Long
		Deck(Card, 7) = CardOutOfPlay
		
		'UPGRADE_ISSUE: Form method PlayTable.Line was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
        '		Me.Line (Deck(Card, 1), Deck(Card, 3)) - (nWidth - 1, nHeight - 1), System.Drawing.ColorTranslator.ToOle(PlayTable.DefInstance.BackColor), BF
		
	End Sub
	Sub UpdateStatusBar()
		
		StatusBar1.Panels(1).Text = PlayerNames(1) & " : " & PlayerScore(1)
		
		If Player = 1 Then
			StatusBar1.Panels(2).Text = PlayerNames(1) & "'s Turn"
		Else
			StatusBar1.Panels(2).Text = PlayerNames(2) & "'s Turn"
		End If
		
		StatusBar1.Panels(3).Text = PlayerNames(2) & " : " & PlayerScore(2)
		
	End Sub
	Sub UpdateScore()
		
		If CardsMatch Then
			If Player = 1 Then
				PlayerScore(1) = PlayerScore(1) + 2
			Else
				PlayerScore(2) = PlayerScore(2) + 2
			End If
		Else
			If Player = 1 Then
				Player = 2
			Else
				Player = 1
			End If
		End If
		
		Call UpdateStatusBar()
		
		If Not TheGameIsOver Then
			Timer1.Enabled = True
		End If
		
	End Sub
	Sub AllDone()
		
		Dim Ret As Integer
		
		If NewGameRequested Then
			frmOptions.DefInstance.FlipTimeOnMatch.Text = CStr(HoldOptions(1))
			frmOptions.DefInstance.FlipTimeOnNoMatch.Text = CStr(HoldOptions(2))
			frmOptions.DefInstance.FlipTimeBetweenComputerSelections.Text = CStr(HoldOptions(3))
			frmOptions.DefInstance.TimeBetweenComputerMatches.Text = CStr(HoldOptions(4))
		End If
		
		If frmOptions.DefInstance.SaveOptionsOnExitCheck.CheckState = 1 Then
			ManageDefaults(("Write Everything"))
		Else
			ManageDefaults(("Write Position Only"))
		End If
		
		Ret = cdtTerm()
		End
		
	End Sub
	Sub EvalutePlayTable()
		
		Dim ChosenCard(2) As Short
		Dim FoundCard As Short
		Dim Card As Short
		
		FoundCard = 0
		ChosenCard(1) = -1
		ChosenCard(2) = -2
		
		' Find all cards which have been chosen
		For Card = 1 To 52
			If (Deck(Card, 5) = C_FACES) And (Deck(Card, 7) = CardInPlay) Then
				FoundCard = FoundCard + 1
				ChosenCard(FoundCard) = Deck(Card, 6)
			End If
		Next Card
		
		' if two cards have been chosen then
		'   if they have the same value
		'       then wait a period of time and take them both off the board
		'       otherwise wait for a period of time and hide the two cards
		' once one of the two above options above have completed update the score
		
		If FoundCard = 2 Then
			NewGameLockout = True
			If ChosenCard(1) = ChosenCard(2) Then
				Delay(CDbl(frmOptions.DefInstance.FlipTimeOnMatch.Text))
				CardsMatch = True
				For Card = 1 To 52
					If (Deck(Card, 5) = C_FACES) And (Deck(Card, 7) = CardInPlay) Then
						RemoveCard((Card))
					End If
				Next 
			Else
				Delay(CDbl(frmOptions.DefInstance.FlipTimeOnNoMatch.Text))
				CardsMatch = False
				For Card = 1 To 52
					If (Deck(Card, 5) = C_FACES) And (Deck(Card, 7) = CardInPlay) Then
						HideCard((Card))
					End If
				Next 
			End If
			Call UpdateScore()
			NewGameLockout = False
			If NewGameRequested Then
				Call NewGame()
			End If
		End If
		
	End Sub
	Function TheGameIsOver() As Boolean
		
		Dim Card As Short
		
		TheGameIsOver = True
		'if all cards have been taken off the board then the game is over
		
		For Card = 1 To 52
			If Deck(Card, 7) = CardInPlay Then
				TheGameIsOver = False
				Exit For
			End If
		Next 
		
		If TheGameIsOver Then
			StatusBar1.Panels(2).Alignment = ComctlLib.PanelAlignmentConstants.sbrCenter
			StatusBar1.Panels(2).Text = "Please press [F2] if you would like to play another game."
			WinnersBox.Visible = True
			If PlayerScore(1) > PlayerScore(2) Then
				WinnersBox.Text = "Congratulations" & vbCr & PlayerNames(1) & " Wins!"
			Else
				If PlayerScore(2) > PlayerScore(1) Then
					WinnersBox.Text = "Congratulations" & vbCr & PlayerNames(2) & " Wins!"
				Else
					WinnersBox.Text = "A Tie!"
				End If
			End If
			Timer2.Enabled = True
		End If
		
	End Function
	Private Sub PlayTable_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = eventArgs.X
		Dim Y As Single = eventArgs.Y
		
		Dim CardFound As Boolean
		Dim CardCount As Short
		Dim Card As Short
		
		If frmOptions.DefInstance.PlayerComputer(Player).Checked Then
			GoTo EarlyExit
		End If
		
		CardFound = False
		CardCount = 0
		
		' Only look for card if less than two have been selected
		For Card = 1 To 52
			If Deck(Card, 5) = C_FACES Then
				If Deck(Card, 7) = CardInPlay Then
					CardCount = CardCount + 1
				End If
			End If
		Next 
		
		If CardCount < 2 Then
			Call LocateCard(X, Y, Card, CardFound)
		End If
		
		If CardFound Then
			Call ShowCard(Card)
			Call EvalutePlayTable()
		End If
		
EarlyExit: 
		
	End Sub
	
	Public Sub MenuGameDeal_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MenuGameDeal.Popup
		MenuGameDeal_Click(eventSender, eventArgs)
	End Sub
	Public Sub MenuGameDeal_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MenuGameDeal.Click
		
		Call NewGame()
		
	End Sub
	Public Sub MenuGameExit_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MenuGameExit.Popup
		MenuGameExit_Click(eventSender, eventArgs)
	End Sub
	Public Sub MenuGameExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MenuGameExit.Click
		
		Call AllDone()
		
	End Sub
	Public Sub MenuGameOptions_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MenuGameOptions.Popup
		MenuGameOptions_Click(eventSender, eventArgs)
	End Sub
	Public Sub MenuGameOptions_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MenuGameOptions.Click
		
		'UPGRADE_NOTE: Modal was upgraded to Modal_Renamed. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1061"'
		Const Modal_Renamed As Short = 1
		
		GameInProgress = False
		Timer1.Enabled = False
		
		VB6.ShowForm(frmOptions.DefInstance, Modal_Renamed)
		
	End Sub
	Public Sub MenuHelpAbout_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MenuHelpAbout.Popup
		MenuHelpAbout_Click(eventSender, eventArgs)
	End Sub
	Public Sub MenuHelpAbout_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MenuHelpAbout.Click
		
		frmAbout.DefInstance.Show()
		
	End Sub
	
	Public Sub MenuHelpHelpTopics_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MenuHelpHelpTopics.Popup
		MenuHelpHelpTopics_Click(eventSender, eventArgs)
	End Sub
	Public Sub MenuHelpHelpTopics_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MenuHelpHelpTopics.Click
		
		CommonDialog1.HelpFile = VB6.GetPath & "\Concentration.HLP"
		CommonDialog1.HelpCommand = MSComDlg.HelpConstants.cdlHelpContents
		CommonDialog1.ShowHelp()
		
	End Sub
	
	Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
		
		Dim Card As Short
		
		Timer1.Enabled = False
		GameInProgress = False
		
		For Card = 1 To 52
			If Deck(Card, 7) = CardInPlay Then
				GameInProgress = True
				Exit For
			End If
		Next Card
		
		If GameInProgress Then
			If frmOptions.DefInstance.PlayerComputer(Player).Checked Then
				Call ComputerPlay()
			End If
		End If
		
	End Sub
	Private Sub Timer2_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer2.Tick
		
		
		If (frmOptions.DefInstance.PlayerComputer(1).Checked = True) And (frmOptions.DefInstance.PlayerComputer(1).Checked = True) Then
			
			Delay(CDbl(frmOptions.DefInstance.TimeBetweenComputerMatches.Text))
			
			If Timer2.Enabled Then
				Call NewGame()
				Timer2.Enabled = False
			End If
			
		End If
		
	End Sub
	Private Sub PlayTable_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		Dim Ret As Integer
		
		If (UBound(Diagnostics.Process.GetProcessesByName(Diagnostics.Process.GetCurrentProcess.ProcessName)) > 0) Then
			Ret = MsgBox(System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location).ProductName & " is already running in the system;" & vbCrLf & "a second copy can not be run simultaneously.", MsgBoxStyle.Exclamation, "Warning")
			End
		End If
		
		CardsMatch = False
		Timer1.Enabled = False
		Timer2.Enabled = False
		
		NewGameLockout = False
		NewGameRequested = False
		
		Ret = cdtInit(nWidth, nHeight)
		OriginalRatio = nHeight / nWidth
		
		StatusBar1.Panels(1).Alignment = ComctlLib.PanelAlignmentConstants.sbrCenter
		StatusBar1.Panels(2).Alignment = ComctlLib.PanelAlignmentConstants.sbrCenter
		StatusBar1.Panels(3).Alignment = ComctlLib.PanelAlignmentConstants.sbrCenter
		
		'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1039"'
        'Load(frmOptions)
		'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1039"'
        'frmAbout
		
		Call NewGame()
		
		PlayTable.DefInstance.Show()
		
	End Sub
	Private Sub PlayTable_Paint(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
		
		Call DrawPlayTable("NoClear")
		
	End Sub
	'UPGRADE_WARNING: Form event PlayTable.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2065"'
	Private Sub PlayTable_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		Dim Card As Short
		
		Call DrawPlayTable("NoClear")
		
		GameInProgress = True
		Timer1.Enabled = True
		
	End Sub
	Private Sub PlayTable_GotFocus(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.GotFocus
		
		Call UpdateStatusBar()
		GameInProgress = True
		Timer1.Enabled = True
		
	End Sub
	'UPGRADE_WARNING: Event PlayTable.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2075"'
	Private Sub PlayTable_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		
		Call DrawPlayTable("Clear")
		
	End Sub
	'UPGRADE_WARNING: Form event PlayTable.Unload has a new behavior. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2065"'
	Private Sub PlayTable_Closed(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Closed
		
		Call AllDone()
		
	End Sub
	'UPGRADE_NOTE: Form_Terminate was upgraded to Form_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1061"'
	'UPGRADE_WARNING: PlayTable event Form.Terminate has a new behavior. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2065"'
	Private Sub Form_Terminate_Renamed()
		
		Call MenuGameExit_Click(MenuGameExit, New System.EventArgs())
		
	End Sub
End Class