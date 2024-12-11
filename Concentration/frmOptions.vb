Option Strict Off
Option Explicit On
Friend Class frmOptions
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
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents SpinMatch As AxComCtl2.AxUpDown
	Public WithEvents _SpinNoMatch_0 As AxComCtl2.AxUpDown
	Public WithEvents _SpinNoMatch_1 As AxComCtl2.AxUpDown
	Public WithEvents _SpinNoMatch_2 As AxComCtl2.AxUpDown
	Public WithEvents TimeBetweenComputerMatches As System.Windows.Forms.Label
	Public WithEvents _Label1_3 As System.Windows.Forms.Label
	Public WithEvents _Label1_2 As System.Windows.Forms.Label
	Public WithEvents FlipTimeBetweenComputerSelections As System.Windows.Forms.Label
	Public WithEvents _Label1_0 As System.Windows.Forms.Label
	Public WithEvents _Label1_1 As System.Windows.Forms.Label
	Public WithEvents FlipTimeOnMatch As System.Windows.Forms.Label
	Public WithEvents FlipTimeOnNoMatch As System.Windows.Forms.Label
	Public WithEvents FlipTimeFrame As System.Windows.Forms.GroupBox
	Public WithEvents CmdCancel As System.Windows.Forms.Button
	Public WithEvents CmdOK As System.Windows.Forms.Button
	Public WithEvents TabStrip1 As AxComctlLib.AxTabStrip
	Public WithEvents _PlayerComputer_2 As System.Windows.Forms.RadioButton
	Public WithEvents _PlayerPerson_2 As System.Windows.Forms.RadioButton
	Public WithEvents _PlayerName_2 As System.Windows.Forms.TextBox
	Public WithEvents _Slider_2 As AxComctlLib.AxSlider
	Public WithEvents _lblS2_2 As System.Windows.Forms.Label
	Public WithEvents _lblS1_2 As System.Windows.Forms.Label
	Public WithEvents Player2Frame As System.Windows.Forms.GroupBox
	Public WithEvents SaveOptionsOnExitCheck As System.Windows.Forms.CheckBox
	Public WithEvents OnExitFrame As System.Windows.Forms.GroupBox
	Public WithEvents _PlayerName_1 As System.Windows.Forms.TextBox
	Public WithEvents _PlayerPerson_1 As System.Windows.Forms.RadioButton
	Public WithEvents _PlayerComputer_1 As System.Windows.Forms.RadioButton
	Public WithEvents _Slider_1 As AxComctlLib.AxSlider
	Public WithEvents _lblS1_1 As System.Windows.Forms.Label
	Public WithEvents _lblS2_1 As System.Windows.Forms.Label
	Public WithEvents Player1Frame As System.Windows.Forms.GroupBox
	Public WithEvents lblBackOfCardName As System.Windows.Forms.Label
	Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents PlayerComputer As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
	Public WithEvents PlayerName As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
	Public WithEvents PlayerPerson As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
	Public WithEvents Slider As AxSliderArray.AxSliderArray
	Public WithEvents SpinNoMatch As AxUpDownArray.AxUpDownArray
	Public WithEvents lblS1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents lblS2 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmOptions))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.FlipTimeFrame = New System.Windows.Forms.GroupBox
        Me.SpinMatch = New AxComCtl2.AxUpDown
        Me._SpinNoMatch_0 = New AxComCtl2.AxUpDown
        Me._SpinNoMatch_1 = New AxComCtl2.AxUpDown
        Me._SpinNoMatch_2 = New AxComCtl2.AxUpDown
        Me.TimeBetweenComputerMatches = New System.Windows.Forms.Label
        Me._Label1_3 = New System.Windows.Forms.Label
        Me._Label1_2 = New System.Windows.Forms.Label
        Me.FlipTimeBetweenComputerSelections = New System.Windows.Forms.Label
        Me._Label1_0 = New System.Windows.Forms.Label
        Me._Label1_1 = New System.Windows.Forms.Label
        Me.FlipTimeOnMatch = New System.Windows.Forms.Label
        Me.FlipTimeOnNoMatch = New System.Windows.Forms.Label
        Me.CmdCancel = New System.Windows.Forms.Button
        Me.CmdOK = New System.Windows.Forms.Button
        Me.TabStrip1 = New AxComctlLib.AxTabStrip
        Me.Player2Frame = New System.Windows.Forms.GroupBox
        Me._PlayerComputer_2 = New System.Windows.Forms.RadioButton
        Me._PlayerPerson_2 = New System.Windows.Forms.RadioButton
        Me._PlayerName_2 = New System.Windows.Forms.TextBox
        Me._Slider_2 = New AxComctlLib.AxSlider
        Me._lblS2_2 = New System.Windows.Forms.Label
        Me._lblS1_2 = New System.Windows.Forms.Label
        Me.OnExitFrame = New System.Windows.Forms.GroupBox
        Me.SaveOptionsOnExitCheck = New System.Windows.Forms.CheckBox
        Me.Player1Frame = New System.Windows.Forms.GroupBox
        Me._PlayerName_1 = New System.Windows.Forms.TextBox
        Me._PlayerPerson_1 = New System.Windows.Forms.RadioButton
        Me._PlayerComputer_1 = New System.Windows.Forms.RadioButton
        Me._Slider_1 = New AxComctlLib.AxSlider
        Me._lblS1_1 = New System.Windows.Forms.Label
        Me._lblS2_1 = New System.Windows.Forms.Label
        Me.lblBackOfCardName = New System.Windows.Forms.Label
        Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.PlayerComputer = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me.PlayerName = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
        Me.PlayerPerson = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me.Slider = New AxSliderArray.AxSliderArray(Me.components)
        Me.SpinNoMatch = New AxUpDownArray.AxUpDownArray(Me.components)
        Me.lblS1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblS2 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.FlipTimeFrame.SuspendLayout()
        CType(Me.SpinMatch, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._SpinNoMatch_0, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._SpinNoMatch_1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._SpinNoMatch_2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TabStrip1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Player2Frame.SuspendLayout()
        CType(Me._Slider_2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.OnExitFrame.SuspendLayout()
        Me.Player1Frame.SuspendLayout()
        CType(Me._Slider_1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PlayerComputer, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PlayerName, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PlayerPerson, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Slider, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SpinNoMatch, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblS1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblS2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'FlipTimeFrame
        '
        Me.FlipTimeFrame.BackColor = System.Drawing.SystemColors.Control
        Me.FlipTimeFrame.Controls.Add(Me.SpinMatch)
        Me.FlipTimeFrame.Controls.Add(Me._SpinNoMatch_0)
        Me.FlipTimeFrame.Controls.Add(Me._SpinNoMatch_1)
        Me.FlipTimeFrame.Controls.Add(Me._SpinNoMatch_2)
        Me.FlipTimeFrame.Controls.Add(Me.TimeBetweenComputerMatches)
        Me.FlipTimeFrame.Controls.Add(Me._Label1_3)
        Me.FlipTimeFrame.Controls.Add(Me._Label1_2)
        Me.FlipTimeFrame.Controls.Add(Me.FlipTimeBetweenComputerSelections)
        Me.FlipTimeFrame.Controls.Add(Me._Label1_0)
        Me.FlipTimeFrame.Controls.Add(Me._Label1_1)
        Me.FlipTimeFrame.Controls.Add(Me.FlipTimeOnMatch)
        Me.FlipTimeFrame.Controls.Add(Me.FlipTimeOnNoMatch)
        Me.FlipTimeFrame.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FlipTimeFrame.ForeColor = System.Drawing.SystemColors.ControlText
        Me.FlipTimeFrame.Location = New System.Drawing.Point(6, 29)
        Me.FlipTimeFrame.Name = "FlipTimeFrame"
        Me.FlipTimeFrame.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.FlipTimeFrame.Size = New System.Drawing.Size(454, 181)
        Me.FlipTimeFrame.TabIndex = 4
        Me.FlipTimeFrame.TabStop = False
        Me.FlipTimeFrame.Text = "Delay time in seconds"
        '
        'SpinMatch
        '
        Me.SpinMatch.ContainingControl = Me
        Me.SpinMatch.Location = New System.Drawing.Point(212, 18)
        Me.SpinMatch.Name = "SpinMatch"
        Me.SpinMatch.OcxState = CType(resources.GetObject("SpinMatch.OcxState"), System.Windows.Forms.AxHost.State)
        Me.SpinMatch.Size = New System.Drawing.Size(17, 30)
        Me.SpinMatch.TabIndex = 19
        '
        '_SpinNoMatch_0
        '
        Me._SpinNoMatch_0.ContainingControl = Me
        Me.SpinNoMatch.SetIndex(Me._SpinNoMatch_0, CType(0, Short))
        Me._SpinNoMatch_0.Location = New System.Drawing.Point(213, 58)
        Me._SpinNoMatch_0.Name = "_SpinNoMatch_0"
        Me._SpinNoMatch_0.OcxState = CType(resources.GetObject("_SpinNoMatch_0.OcxState"), System.Windows.Forms.AxHost.State)
        Me._SpinNoMatch_0.Size = New System.Drawing.Size(17, 30)
        Me._SpinNoMatch_0.TabIndex = 20
        '
        '_SpinNoMatch_1
        '
        Me._SpinNoMatch_1.ContainingControl = Me
        Me.SpinNoMatch.SetIndex(Me._SpinNoMatch_1, CType(1, Short))
        Me._SpinNoMatch_1.Location = New System.Drawing.Point(212, 98)
        Me._SpinNoMatch_1.Name = "_SpinNoMatch_1"
        Me._SpinNoMatch_1.OcxState = CType(resources.GetObject("_SpinNoMatch_1.OcxState"), System.Windows.Forms.AxHost.State)
        Me._SpinNoMatch_1.Size = New System.Drawing.Size(17, 28)
        Me._SpinNoMatch_1.TabIndex = 25
        '
        '_SpinNoMatch_2
        '
        Me._SpinNoMatch_2.ContainingControl = Me
        Me.SpinNoMatch.SetIndex(Me._SpinNoMatch_2, CType(2, Short))
        Me._SpinNoMatch_2.Location = New System.Drawing.Point(213, 136)
        Me._SpinNoMatch_2.Name = "_SpinNoMatch_2"
        Me._SpinNoMatch_2.OcxState = CType(resources.GetObject("_SpinNoMatch_2.OcxState"), System.Windows.Forms.AxHost.State)
        Me._SpinNoMatch_2.Size = New System.Drawing.Size(17, 30)
        Me._SpinNoMatch_2.TabIndex = 30
        '
        'TimeBetweenComputerMatches
        '
        Me.TimeBetweenComputerMatches.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.TimeBetweenComputerMatches.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.TimeBetweenComputerMatches.Cursor = System.Windows.Forms.Cursors.Default
        Me.TimeBetweenComputerMatches.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TimeBetweenComputerMatches.ForeColor = System.Drawing.SystemColors.ControlText
        Me.TimeBetweenComputerMatches.Location = New System.Drawing.Point(190, 140)
        Me.TimeBetweenComputerMatches.Name = "TimeBetweenComputerMatches"
        Me.TimeBetweenComputerMatches.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.TimeBetweenComputerMatches.Size = New System.Drawing.Size(23, 15)
        Me.TimeBetweenComputerMatches.TabIndex = 32
        Me.TimeBetweenComputerMatches.Text = "2"
        Me.TimeBetweenComputerMatches.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_Label1_3
        '
        Me._Label1_3.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.SetIndex(Me._Label1_3, CType(3, Short))
        Me._Label1_3.Location = New System.Drawing.Point(33, 143)
        Me._Label1_3.Name = "_Label1_3"
        Me._Label1_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_3.Size = New System.Drawing.Size(149, 18)
        Me._Label1_3.TabIndex = 31
        Me._Label1_3.Text = "Between computer matches"
        '
        '_Label1_2
        '
        Me._Label1_2.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.SetIndex(Me._Label1_2, CType(2, Short))
        Me._Label1_2.Location = New System.Drawing.Point(33, 105)
        Me._Label1_2.Name = "_Label1_2"
        Me._Label1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_2.Size = New System.Drawing.Size(149, 18)
        Me._Label1_2.TabIndex = 27
        Me._Label1_2.Text = "Between computer selections"
        '
        'FlipTimeBetweenComputerSelections
        '
        Me.FlipTimeBetweenComputerSelections.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.FlipTimeBetweenComputerSelections.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.FlipTimeBetweenComputerSelections.Cursor = System.Windows.Forms.Cursors.Default
        Me.FlipTimeBetweenComputerSelections.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FlipTimeBetweenComputerSelections.ForeColor = System.Drawing.SystemColors.ControlText
        Me.FlipTimeBetweenComputerSelections.Location = New System.Drawing.Point(190, 102)
        Me.FlipTimeBetweenComputerSelections.Name = "FlipTimeBetweenComputerSelections"
        Me.FlipTimeBetweenComputerSelections.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.FlipTimeBetweenComputerSelections.Size = New System.Drawing.Size(23, 15)
        Me.FlipTimeBetweenComputerSelections.TabIndex = 26
        Me.FlipTimeBetweenComputerSelections.Text = "2"
        Me.FlipTimeBetweenComputerSelections.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_Label1_0
        '
        Me._Label1_0.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.SetIndex(Me._Label1_0, CType(0, Short))
        Me._Label1_0.Location = New System.Drawing.Point(33, 27)
        Me._Label1_0.Name = "_Label1_0"
        Me._Label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_0.Size = New System.Drawing.Size(107, 18)
        Me._Label1_0.TabIndex = 24
        Me._Label1_0.Text = "When there is a match"
        '
        '_Label1_1
        '
        Me._Label1_1.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.SetIndex(Me._Label1_1, CType(1, Short))
        Me._Label1_1.Location = New System.Drawing.Point(33, 66)
        Me._Label1_1.Name = "_Label1_1"
        Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_1.Size = New System.Drawing.Size(135, 18)
        Me._Label1_1.TabIndex = 23
        Me._Label1_1.Text = "When there is no match"
        '
        'FlipTimeOnMatch
        '
        Me.FlipTimeOnMatch.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.FlipTimeOnMatch.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.FlipTimeOnMatch.Cursor = System.Windows.Forms.Cursors.Default
        Me.FlipTimeOnMatch.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FlipTimeOnMatch.ForeColor = System.Drawing.SystemColors.ControlText
        Me.FlipTimeOnMatch.Location = New System.Drawing.Point(190, 24)
        Me.FlipTimeOnMatch.Name = "FlipTimeOnMatch"
        Me.FlipTimeOnMatch.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.FlipTimeOnMatch.Size = New System.Drawing.Size(23, 15)
        Me.FlipTimeOnMatch.TabIndex = 22
        Me.FlipTimeOnMatch.Text = "2"
        Me.FlipTimeOnMatch.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'FlipTimeOnNoMatch
        '
        Me.FlipTimeOnNoMatch.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.FlipTimeOnNoMatch.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.FlipTimeOnNoMatch.Cursor = System.Windows.Forms.Cursors.Default
        Me.FlipTimeOnNoMatch.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FlipTimeOnNoMatch.ForeColor = System.Drawing.SystemColors.ControlText
        Me.FlipTimeOnNoMatch.Location = New System.Drawing.Point(190, 63)
        Me.FlipTimeOnNoMatch.Name = "FlipTimeOnNoMatch"
        Me.FlipTimeOnNoMatch.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.FlipTimeOnNoMatch.Size = New System.Drawing.Size(23, 15)
        Me.FlipTimeOnNoMatch.TabIndex = 21
        Me.FlipTimeOnNoMatch.Text = "3"
        Me.FlipTimeOnNoMatch.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'CmdCancel
        '
        Me.CmdCancel.BackColor = System.Drawing.SystemColors.Control
        Me.CmdCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.CmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.CmdCancel.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.CmdCancel.Location = New System.Drawing.Point(239, 243)
        Me.CmdCancel.Name = "CmdCancel"
        Me.CmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.CmdCancel.Size = New System.Drawing.Size(224, 33)
        Me.CmdCancel.TabIndex = 2
        Me.CmdCancel.Text = "Cancel"
        '
        'CmdOK
        '
        Me.CmdOK.BackColor = System.Drawing.SystemColors.Control
        Me.CmdOK.Cursor = System.Windows.Forms.Cursors.Default
        Me.CmdOK.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmdOK.ForeColor = System.Drawing.SystemColors.ControlText
        Me.CmdOK.Location = New System.Drawing.Point(7, 243)
        Me.CmdOK.Name = "CmdOK"
        Me.CmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.CmdOK.Size = New System.Drawing.Size(225, 33)
        Me.CmdOK.TabIndex = 0
        Me.CmdOK.Text = "Ok"
        '
        'TabStrip1
        '
        Me.TabStrip1.Location = New System.Drawing.Point(0, 0)
        Me.TabStrip1.Name = "TabStrip1"
        Me.TabStrip1.OcxState = CType(resources.GetObject("TabStrip1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.TabStrip1.Size = New System.Drawing.Size(473, 21)
        Me.TabStrip1.TabIndex = 1
        '
        'Player2Frame
        '
        Me.Player2Frame.BackColor = System.Drawing.SystemColors.Control
        Me.Player2Frame.Controls.Add(Me._PlayerComputer_2)
        Me.Player2Frame.Controls.Add(Me._PlayerPerson_2)
        Me.Player2Frame.Controls.Add(Me._PlayerName_2)
        Me.Player2Frame.Controls.Add(Me._Slider_2)
        Me.Player2Frame.Controls.Add(Me._lblS2_2)
        Me.Player2Frame.Controls.Add(Me._lblS1_2)
        Me.Player2Frame.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Player2Frame.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Player2Frame.Location = New System.Drawing.Point(237, 56)
        Me.Player2Frame.Name = "Player2Frame"
        Me.Player2Frame.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Player2Frame.Size = New System.Drawing.Size(211, 104)
        Me.Player2Frame.TabIndex = 12
        Me.Player2Frame.TabStop = False
        Me.Player2Frame.Text = "Player Two"
        '
        '_PlayerComputer_2
        '
        Me._PlayerComputer_2.BackColor = System.Drawing.SystemColors.Control
        Me._PlayerComputer_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._PlayerComputer_2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._PlayerComputer_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.PlayerComputer.SetIndex(Me._PlayerComputer_2, CType(2, Short))
        Me._PlayerComputer_2.Location = New System.Drawing.Point(12, 50)
        Me._PlayerComputer_2.Name = "_PlayerComputer_2"
        Me._PlayerComputer_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._PlayerComputer_2.Size = New System.Drawing.Size(69, 19)
        Me._PlayerComputer_2.TabIndex = 15
        Me._PlayerComputer_2.TabStop = True
        Me._PlayerComputer_2.Text = "Computer"
        '
        '_PlayerPerson_2
        '
        Me._PlayerPerson_2.BackColor = System.Drawing.SystemColors.Control
        Me._PlayerPerson_2.Checked = True
        Me._PlayerPerson_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._PlayerPerson_2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._PlayerPerson_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.PlayerPerson.SetIndex(Me._PlayerPerson_2, CType(2, Short))
        Me._PlayerPerson_2.Location = New System.Drawing.Point(12, 26)
        Me._PlayerPerson_2.Name = "_PlayerPerson_2"
        Me._PlayerPerson_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._PlayerPerson_2.Size = New System.Drawing.Size(58, 20)
        Me._PlayerPerson_2.TabIndex = 14
        Me._PlayerPerson_2.TabStop = True
        Me._PlayerPerson_2.Text = "Person"
        '
        '_PlayerName_2
        '
        Me._PlayerName_2.AcceptsReturn = True
        Me._PlayerName_2.AutoSize = False
        Me._PlayerName_2.BackColor = System.Drawing.SystemColors.Window
        Me._PlayerName_2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._PlayerName_2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._PlayerName_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.PlayerName.SetIndex(Me._PlayerName_2, CType(2, Short))
        Me._PlayerName_2.Location = New System.Drawing.Point(74, 28)
        Me._PlayerName_2.MaxLength = 0
        Me._PlayerName_2.Name = "_PlayerName_2"
        Me._PlayerName_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._PlayerName_2.Size = New System.Drawing.Size(103, 20)
        Me._PlayerName_2.TabIndex = 13
        Me._PlayerName_2.Text = "Player Two"
        '
        '_Slider_2
        '
        Me._Slider_2.ContainingControl = Me
        Me.Slider.SetIndex(Me._Slider_2, CType(2, Short))
        Me._Slider_2.Location = New System.Drawing.Point(69, 76)
        Me._Slider_2.Name = "_Slider_2"
        Me._Slider_2.OcxState = CType(resources.GetObject("_Slider_2.OcxState"), System.Windows.Forms.AxHost.State)
        Me._Slider_2.Size = New System.Drawing.Size(86, 16)
        Me._Slider_2.TabIndex = 16
        '
        '_lblS2_2
        '
        Me._lblS2_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblS2_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblS2_2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblS2_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblS2.SetIndex(Me._lblS2_2, CType(2, Short))
        Me._lblS2_2.Location = New System.Drawing.Point(157, 77)
        Me._lblS2_2.Name = "_lblS2_2"
        Me._lblS2_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblS2_2.Size = New System.Drawing.Size(36, 12)
        Me._lblS2_2.TabIndex = 18
        Me._lblS2_2.Text = "Expert"
        '
        '_lblS1_2
        '
        Me._lblS1_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblS1_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblS1_2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblS1_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblS1.SetIndex(Me._lblS1_2, CType(2, Short))
        Me._lblS1_2.Location = New System.Drawing.Point(32, 76)
        Me._lblS1_2.Name = "_lblS1_2"
        Me._lblS1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblS1_2.Size = New System.Drawing.Size(47, 12)
        Me._lblS1_2.TabIndex = 17
        Me._lblS1_2.Text = "Novice"
        '
        'OnExitFrame
        '
        Me.OnExitFrame.BackColor = System.Drawing.SystemColors.Control
        Me.OnExitFrame.Controls.Add(Me.SaveOptionsOnExitCheck)
        Me.OnExitFrame.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.OnExitFrame.ForeColor = System.Drawing.SystemColors.ControlText
        Me.OnExitFrame.Location = New System.Drawing.Point(5, 57)
        Me.OnExitFrame.Name = "OnExitFrame"
        Me.OnExitFrame.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.OnExitFrame.Size = New System.Drawing.Size(453, 134)
        Me.OnExitFrame.TabIndex = 28
        Me.OnExitFrame.TabStop = False
        Me.OnExitFrame.Text = "On Exit"
        '
        'SaveOptionsOnExitCheck
        '
        Me.SaveOptionsOnExitCheck.BackColor = System.Drawing.SystemColors.Control
        Me.SaveOptionsOnExitCheck.Cursor = System.Windows.Forms.Cursors.Default
        Me.SaveOptionsOnExitCheck.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SaveOptionsOnExitCheck.ForeColor = System.Drawing.SystemColors.ControlText
        Me.SaveOptionsOnExitCheck.Location = New System.Drawing.Point(25, 34)
        Me.SaveOptionsOnExitCheck.Name = "SaveOptionsOnExitCheck"
        Me.SaveOptionsOnExitCheck.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.SaveOptionsOnExitCheck.Size = New System.Drawing.Size(165, 19)
        Me.SaveOptionsOnExitCheck.TabIndex = 29
        Me.SaveOptionsOnExitCheck.Text = "Remember Options on Exit"
        '
        'Player1Frame
        '
        Me.Player1Frame.BackColor = System.Drawing.SystemColors.Control
        Me.Player1Frame.Controls.Add(Me._PlayerName_1)
        Me.Player1Frame.Controls.Add(Me._PlayerPerson_1)
        Me.Player1Frame.Controls.Add(Me._PlayerComputer_1)
        Me.Player1Frame.Controls.Add(Me._Slider_1)
        Me.Player1Frame.Controls.Add(Me._lblS1_1)
        Me.Player1Frame.Controls.Add(Me._lblS2_1)
        Me.Player1Frame.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Player1Frame.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Player1Frame.Location = New System.Drawing.Point(15, 56)
        Me.Player1Frame.Name = "Player1Frame"
        Me.Player1Frame.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Player1Frame.Size = New System.Drawing.Size(211, 104)
        Me.Player1Frame.TabIndex = 5
        Me.Player1Frame.TabStop = False
        Me.Player1Frame.Text = "Player One"
        '
        '_PlayerName_1
        '
        Me._PlayerName_1.AcceptsReturn = True
        Me._PlayerName_1.AutoSize = False
        Me._PlayerName_1.BackColor = System.Drawing.SystemColors.Window
        Me._PlayerName_1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._PlayerName_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._PlayerName_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.PlayerName.SetIndex(Me._PlayerName_1, CType(1, Short))
        Me._PlayerName_1.Location = New System.Drawing.Point(74, 28)
        Me._PlayerName_1.MaxLength = 0
        Me._PlayerName_1.Name = "_PlayerName_1"
        Me._PlayerName_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._PlayerName_1.Size = New System.Drawing.Size(103, 20)
        Me._PlayerName_1.TabIndex = 8
        Me._PlayerName_1.Text = "Player One"
        '
        '_PlayerPerson_1
        '
        Me._PlayerPerson_1.BackColor = System.Drawing.SystemColors.Control
        Me._PlayerPerson_1.Checked = True
        Me._PlayerPerson_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._PlayerPerson_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._PlayerPerson_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.PlayerPerson.SetIndex(Me._PlayerPerson_1, CType(1, Short))
        Me._PlayerPerson_1.Location = New System.Drawing.Point(12, 26)
        Me._PlayerPerson_1.Name = "_PlayerPerson_1"
        Me._PlayerPerson_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._PlayerPerson_1.Size = New System.Drawing.Size(58, 20)
        Me._PlayerPerson_1.TabIndex = 7
        Me._PlayerPerson_1.TabStop = True
        Me._PlayerPerson_1.Text = "Person"
        '
        '_PlayerComputer_1
        '
        Me._PlayerComputer_1.BackColor = System.Drawing.SystemColors.Control
        Me._PlayerComputer_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._PlayerComputer_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._PlayerComputer_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.PlayerComputer.SetIndex(Me._PlayerComputer_1, CType(1, Short))
        Me._PlayerComputer_1.Location = New System.Drawing.Point(12, 50)
        Me._PlayerComputer_1.Name = "_PlayerComputer_1"
        Me._PlayerComputer_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._PlayerComputer_1.Size = New System.Drawing.Size(69, 19)
        Me._PlayerComputer_1.TabIndex = 6
        Me._PlayerComputer_1.TabStop = True
        Me._PlayerComputer_1.Text = "Computer"
        '
        '_Slider_1
        '
        Me._Slider_1.ContainingControl = Me
        Me.Slider.SetIndex(Me._Slider_1, CType(1, Short))
        Me._Slider_1.Location = New System.Drawing.Point(69, 76)
        Me._Slider_1.Name = "_Slider_1"
        Me._Slider_1.OcxState = CType(resources.GetObject("_Slider_1.OcxState"), System.Windows.Forms.AxHost.State)
        Me._Slider_1.Size = New System.Drawing.Size(86, 16)
        Me._Slider_1.TabIndex = 9
        '
        '_lblS1_1
        '
        Me._lblS1_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblS1_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblS1_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblS1_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblS1.SetIndex(Me._lblS1_1, CType(1, Short))
        Me._lblS1_1.Location = New System.Drawing.Point(32, 76)
        Me._lblS1_1.Name = "_lblS1_1"
        Me._lblS1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblS1_1.Size = New System.Drawing.Size(47, 12)
        Me._lblS1_1.TabIndex = 11
        Me._lblS1_1.Text = "Novice"
        '
        '_lblS2_1
        '
        Me._lblS2_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblS2_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblS2_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblS2_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblS2.SetIndex(Me._lblS2_1, CType(1, Short))
        Me._lblS2_1.Location = New System.Drawing.Point(157, 77)
        Me._lblS2_1.Name = "_lblS2_1"
        Me._lblS2_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblS2_1.Size = New System.Drawing.Size(36, 12)
        Me._lblS2_1.TabIndex = 10
        Me._lblS2_1.Text = "Expert"
        '
        'lblBackOfCardName
        '
        Me.lblBackOfCardName.BackColor = System.Drawing.SystemColors.Control
        Me.lblBackOfCardName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblBackOfCardName.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblBackOfCardName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBackOfCardName.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblBackOfCardName.Location = New System.Drawing.Point(6, 218)
        Me.lblBackOfCardName.Name = "lblBackOfCardName"
        Me.lblBackOfCardName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblBackOfCardName.Size = New System.Drawing.Size(455, 18)
        Me.lblBackOfCardName.TabIndex = 3
        Me.lblBackOfCardName.Text = "Island"
        Me.lblBackOfCardName.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'PlayerComputer
        '
        '
        'PlayerPerson
        '
        '
        'frmOptions
        '
        Me.AcceptButton = Me.CmdOK
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.CancelButton = Me.CmdCancel
        Me.ClientSize = New System.Drawing.Size(465, 283)
        Me.Controls.Add(Me.FlipTimeFrame)
        Me.Controls.Add(Me.CmdCancel)
        Me.Controls.Add(Me.CmdOK)
        Me.Controls.Add(Me.TabStrip1)
        Me.Controls.Add(Me.Player2Frame)
        Me.Controls.Add(Me.OnExitFrame)
        Me.Controls.Add(Me.Player1Frame)
        Me.Controls.Add(Me.lblBackOfCardName)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(75, 115)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmOptions"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Concentration - Options"
        Me.FlipTimeFrame.ResumeLayout(False)
        CType(Me.SpinMatch, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._SpinNoMatch_0, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._SpinNoMatch_1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._SpinNoMatch_2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TabStrip1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Player2Frame.ResumeLayout(False)
        CType(Me._Slider_2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.OnExitFrame.ResumeLayout(False)
        Me.Player1Frame.ResumeLayout(False)
        CType(Me._Slider_1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PlayerComputer, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PlayerName, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PlayerPerson, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Slider, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SpinNoMatch, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblS1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblS2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As frmOptions
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As frmOptions
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New frmOptions()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 
	' Copyright Rob Latour 1998
	' February 21st, 1998
	
	Const NumberOfDeckSelections As Short = 14
	Dim DeckSelection(14, 5) As Object
	' (x,0) ; deck selection value 53 .. 65
	' (x,1) ; card's min x position            min x & y -> +----+
	' (x,2) ; card's max x position                         !    !
	' (x,3) ; card's min y position                         !    !
	' (x,4) ; card's max y position                         +----+ <- max x & y
	' (x,5) ; card's name
	
	Private Sub InitalOptions(ByVal sFunction As String)
		
		Static Options(17) As Object
		Static Index As Short
		Static X As Short
		
		Select Case sFunction
			Case Is = "Save"
				'UPGRADE_WARNING: Couldn't resolve default property of object Options(0). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				Options(0) = PlayerPerson(1).Checked
				'UPGRADE_WARNING: Couldn't resolve default property of object Options(1). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				Options(1) = PlayerName(1).Text
				'UPGRADE_WARNING: Couldn't resolve default property of object Options(2). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				Options(2) = PlayerComputer(1).Checked
				'UPGRADE_WARNING: Couldn't resolve default property of object Options(3). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				Options(3) = Slider(1).Value
				'UPGRADE_WARNING: Couldn't resolve default property of object Options(4). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				Options(4) = PlayerPerson(2).Checked
				'UPGRADE_WARNING: Couldn't resolve default property of object Options(5). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				Options(5) = PlayerName(2).Text
				'UPGRADE_WARNING: Couldn't resolve default property of object Options(6). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				Options(6) = PlayerComputer(2).Checked
				'UPGRADE_WARNING: Couldn't resolve default property of object Options(7). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				Options(7) = Slider(2).Value
				'UPGRADE_WARNING: Couldn't resolve default property of object Options(8). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				Options(8) = FlipTimeOnMatch.Text
				'UPGRADE_WARNING: Couldn't resolve default property of object Options(9). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				Options(9) = FlipTimeOnNoMatch.Text
				'UPGRADE_WARNING: Couldn't resolve default property of object Options(10). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				Options(10) = FlipTimeBetweenComputerSelections.Text
				'UPGRADE_WARNING: Couldn't resolve default property of object Options(11). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				Options(11) = TimeBetweenComputerMatches.Text
				'UPGRADE_WARNING: Couldn't resolve default property of object Options(12). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				Options(12) = lblBackOfCardName.Text
				For X = 1 To 14
					'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(X, 5). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Options(12). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
					If Options(12) = DeckSelection(X, 5) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(X, 0). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
						SelectedBack = DeckSelection(X, 0)
						Exit For
					End If
				Next X
				'UPGRADE_WARNING: Couldn't resolve default property of object Options(13). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				Options(13) = SaveOptionsOnExitCheck.CheckState
				'UPGRADE_WARNING: Couldn't resolve default property of object Options(14). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				Options(14) = VB6.PixelsToTwipsY(PlayTable.DefInstance.Height)
				'UPGRADE_WARNING: Couldn't resolve default property of object Options(15). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				Options(15) = VB6.PixelsToTwipsX(PlayTable.DefInstance.Width)
				'UPGRADE_WARNING: Couldn't resolve default property of object Options(16). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				Options(16) = VB6.PixelsToTwipsY(PlayTable.DefInstance.Top)
				'UPGRADE_WARNING: Couldn't resolve default property of object Options(17). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				Options(17) = VB6.PixelsToTwipsX(PlayTable.DefInstance.Left)
			Case Is = "Restore"
				'UPGRADE_WARNING: Couldn't resolve default property of object Options(0). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				PlayerPerson(1).Checked = Options(0)
				'UPGRADE_WARNING: Couldn't resolve default property of object Options(1). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				PlayerName(1).Text = Options(1)
				'UPGRADE_WARNING: Couldn't resolve default property of object Options(2). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				PlayerComputer(1).Checked = Options(2)
				'UPGRADE_WARNING: Couldn't resolve default property of object Options(3). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				Slider(1).Value = Options(3)
				'UPGRADE_WARNING: Couldn't resolve default property of object Options(4). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				PlayerPerson(2).Checked = Options(4)
				'UPGRADE_WARNING: Couldn't resolve default property of object Options(5). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				PlayerName(2).Text = Options(5)
				'UPGRADE_WARNING: Couldn't resolve default property of object Options(6). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				PlayerComputer(2).Checked = Options(6)
				'UPGRADE_WARNING: Couldn't resolve default property of object Options(7). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				Slider(2).Value = Options(7)
				'UPGRADE_WARNING: Couldn't resolve default property of object Options(8). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				FlipTimeOnMatch.Text = Options(8)
				'UPGRADE_WARNING: Couldn't resolve default property of object Options(9). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				FlipTimeOnNoMatch.Text = Options(9)
				'UPGRADE_WARNING: Couldn't resolve default property of object Options(10). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				FlipTimeBetweenComputerSelections.Text = Options(10)
				'UPGRADE_WARNING: Couldn't resolve default property of object Options(11). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				TimeBetweenComputerMatches.Text = Options(11)
				'UPGRADE_WARNING: Couldn't resolve default property of object Options(12). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				lblBackOfCardName.Text = Options(12)
				For X = 1 To 14
					'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(X, 5). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Options(12). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
					If Options(12) = DeckSelection(X, 5) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(X, 0). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
						SelectedBack = DeckSelection(X, 0)
						Exit For
					End If
				Next X
				'UPGRADE_WARNING: Couldn't resolve default property of object Options(13). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				SaveOptionsOnExitCheck.CheckState = Options(13)
				On Error Resume Next
				'UPGRADE_WARNING: Couldn't resolve default property of object Options(14). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				PlayTable.DefInstance.Height = VB6.TwipsToPixelsY(Options(14))
				'UPGRADE_WARNING: Couldn't resolve default property of object Options(15). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				PlayTable.DefInstance.Width = VB6.TwipsToPixelsX(Options(15))
				'UPGRADE_WARNING: Couldn't resolve default property of object Options(16). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				PlayTable.DefInstance.Top = VB6.TwipsToPixelsY(Options(16))
				'UPGRADE_WARNING: Couldn't resolve default property of object Options(17). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				PlayTable.DefInstance.Left = VB6.TwipsToPixelsX(Options(17))
				On Error GoTo 0
		End Select
		
	End Sub
	Private Sub frmOptions_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		Call PlayerPerson_CheckedChanged(PlayerPerson.Item(1), New System.EventArgs())
		Call PlayerPerson_CheckedChanged(PlayerPerson.Item(2), New System.EventArgs())
		
		'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(1, 0). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		DeckSelection(1, 0) = 53
		'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(1, 5). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		DeckSelection(1, 5) = "Cross Hatch"
		'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(2, 0). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		DeckSelection(2, 0) = 54
		'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(2, 5). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		DeckSelection(2, 5) = "Weave"
		'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(3, 0). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		DeckSelection(3, 0) = 55
		'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(3, 5). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		DeckSelection(3, 5) = "Tweed"
		'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(4, 0). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		DeckSelection(4, 0) = 56
		'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(4, 5). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		DeckSelection(4, 5) = "Robot"
		'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(5, 0). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		DeckSelection(5, 0) = 57
		'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(5, 5). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		DeckSelection(5, 5) = "Rose"
		'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(6, 0). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		DeckSelection(6, 0) = 58
		'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(6, 5). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		DeckSelection(6, 5) = "Vine of the Night"
		'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(7, 0). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		DeckSelection(7, 0) = 59
		'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(7, 5). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		DeckSelection(7, 5) = "Vine of the Day"
		'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(8, 0). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		DeckSelection(8, 0) = 60
		'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(8, 5). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		DeckSelection(8, 5) = "Four Fish"
		'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(9, 0). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		DeckSelection(9, 0) = 61
		'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(9, 5). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		DeckSelection(9, 5) = "Three Fish"
		'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(10, 0). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		DeckSelection(10, 0) = 62
		'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(10, 5). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		DeckSelection(10, 5) = "Shell"
		'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(11, 0). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		DeckSelection(11, 0) = 63
		'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(11, 5). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		DeckSelection(11, 5) = "Castle"
		'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(12, 0). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		DeckSelection(12, 0) = 64
		'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(12, 5). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		DeckSelection(12, 5) = "Island"
		'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(13, 0). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		DeckSelection(13, 0) = 65
		'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(13, 5). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		DeckSelection(13, 5) = "Handfull of Cards"
		'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(14, 0). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		DeckSelection(14, 0) = 68
		'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(14, 5). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		DeckSelection(14, 5) = "The 'O'"
		
		'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(12, 0). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		SelectedBack = DeckSelection(12, 0)
		'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(12, 5). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		lblBackOfCardName.Text = DeckSelection(12, 5)
		Slider(1).Value = 5
		Slider(2).Value = 5
		FlipTimeOnMatch.Text = CStr(2)
		FlipTimeOnNoMatch.Text = CStr(3)
		FlipTimeBetweenComputerSelections.Text = CStr(3)
		TimeBetweenComputerMatches.Text = CStr(15)
		
		Call ManageDefaults("Read")
		Call DefinePlayerNames()
		Call InitalOptions("Save")
		
	End Sub
	'UPGRADE_WARNING: Form event frmOptions.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2065"'
	Private Sub frmOptions_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		
		Const NumberOfRows As Short = 2
		Const NumberOfColumns As Short = 7
		Dim nWidth As Short
		Dim nHeight As Short
		
		Dim X As Short
		Dim Column As Object
		Dim Row As Short
		Dim ColumnOffset As Object
		Dim RowOffset As Short
		
		Call DefinePlayerNames()
		Call InitalOptions("Save")
		
		nWidth = frmOptions.DefInstance.ClientRectangle.Width / NumberOfColumns
		nHeight = frmOptions.DefInstance.ClientRectangle.Height / NumberOfRows / 1.5
		
		X = 1
		'UPGRADE_WARNING: Couldn't resolve default property of object ColumnOffset. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ColumnOffset = 1
		RowOffset = nHeight / 4
		
		For Row = 0 To NumberOfRows - 1
			For Column = 0 To NumberOfColumns - 1
				'UPGRADE_WARNING: Couldn't resolve default property of object Column. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				'UPGRADE_WARNING: Couldn't resolve default property of object ColumnOffset. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(X, 1). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				DeckSelection(X, 1) = ColumnOffset + (Column * nWidth)
				'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(X, 2). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				DeckSelection(X, 2) = nWidth - 1
				'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(X, 3). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				DeckSelection(X, 3) = RowOffset + (Row * nHeight)
				'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(X, 4). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				DeckSelection(X, 4) = nHeight - 1
				X = X + 1
			Next Column
		Next Row
		
		Call TabStrip1_ClickEvent(TabStrip1, New System.EventArgs())
		
	End Sub
	Private Sub frmOptions_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = eventArgs.X
		Dim Y As Single = eventArgs.Y
		
		Dim Card As Short
		
		' (x,0) ; deck selection value 53 .. 65
		' (x,1) ; card's min x position            min x & y -> +----+
		' (x,2) ; card's max x position                         !    !
		' (x,3) ; card's min y position                         !    !
		' (x,4) ; card's max y position                         +----+ <- max x & y
		
		For Card = 1 To NumberOfDeckSelections
			'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(Card, 1). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			If X >= DeckSelection(Card, 1) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(Card, 2). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(Card, 1). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				If X <= DeckSelection(Card, 1) + DeckSelection(Card, 2) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(Card, 3). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
					If Y >= DeckSelection(Card, 3) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(Card, 4). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
						'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(Card, 3). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
						If Y <= DeckSelection(Card, 3) + DeckSelection(Card, 4) Then
							'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(Card, 5). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
							frmOptions.DefInstance.lblBackOfCardName.Text = DeckSelection(Card, 5)
							'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(Card, 0). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
							SelectedBack = DeckSelection(Card, 0)
							Exit For
						End If
					End If
				End If
			End If
		Next 
		
	End Sub
	'UPGRADE_WARNING: Event PlayerPerson.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2075"'
	Private Sub PlayerPerson_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles PlayerPerson.CheckedChanged
		If eventSender.Checked Then
			Dim Index As Short = PlayerPerson.GetIndex(eventSender)
			
			PlayerPerson(Index).Checked = True
			PlayerName(Index).Enabled = True
			PlayerComputer(Index).Checked = False
			lblS1(Index).Enabled = False
			lblS2(Index).Enabled = False
			Slider(Index).Enabled = False
			
		End If
	End Sub
	'UPGRADE_WARNING: Event PlayerComputer.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2075"'
	Private Sub PlayerComputer_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles PlayerComputer.CheckedChanged
		If eventSender.Checked Then
			Dim Index As Short = PlayerComputer.GetIndex(eventSender)
			
			PlayerPerson(Index).Checked = False
			PlayerName(Index).Enabled = False
			PlayerComputer(Index).Checked = True
			lblS1(Index).Enabled = True
			lblS2(Index).Enabled = True
			Slider(Index).Enabled = True
			
		End If
	End Sub
	Sub DrawCardBacks()
		
		Dim X As Short
        Dim Ret As Integer
        Dim hdc As Long
		
		'UPGRADE_ISSUE: Form method frmOptions.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
        'frmOptions.Cls()
		
		For X = 1 To NumberOfDeckSelections
			'Ret = cdtDrawExt(hdc, DeckSelection(X, 1), DeckSelection(X, 3), nWidth, nHeight, DeckSelection(X, 0), C_BACKS, &HFFFF&)
			'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(X, 0). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(X, 4). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(X, 2). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(X, 3). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			'UPGRADE_WARNING: Couldn't resolve default property of object DeckSelection(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			'UPGRADE_ISSUE: Form property frmOptions.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			Ret = cdtDrawExt(hdc, DeckSelection(X, 1), DeckSelection(X, 3), DeckSelection(X, 2), DeckSelection(X, 4), DeckSelection(X, 0), C_BACKS, &HFFFF)
		Next 
		
	End Sub
	Private Sub TabStrip1_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TabStrip1.ClickEvent
		
		Player1Frame.Visible = False
		Player2Frame.Visible = False
		lblBackOfCardName.Visible = False
		FlipTimeFrame.Visible = False
		OnExitFrame.Visible = False
		
		'UPGRADE_ISSUE: Form method frmOptions.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
        'frmOptions.Cls
		Select Case TabStrip1.SelectedItem._ObjectDefault
			Case Is = "Players"
				Player1Frame.Visible = True
				Player2Frame.Visible = True
			Case Is = "Deck"
				lblBackOfCardName.Visible = True
				Call DrawCardBacks()
			Case Is = "Delay Times"
				FlipTimeFrame.Visible = True
			Case Is = "On Exit"
				OnExitFrame.Visible = True
		End Select
		
	End Sub
	Private Sub DefinePlayerNames()
		
		Dim LiteralOneOrTwo(2) As String
		Dim X As Short
		
		LiteralOneOrTwo(1) = "One"
		LiteralOneOrTwo(2) = "Two"
		
		For X = 1 To 2
			frmOptions.DefInstance.PlayerName(X).Text = Trim(frmOptions.DefInstance.PlayerName(X).Text)
			If frmOptions.DefInstance.PlayerName(X).Text = "" Then
				frmOptions.DefInstance.PlayerName(X).Text = "Player " & LiteralOneOrTwo(X)
			End If
			
			If PlayerPerson(X).Checked = True Then
				PlayerNames(X) = frmOptions.DefInstance.PlayerName(X).Text
			Else
				PlayerNames(X) = "Chip"
			End If
		Next X
		
		If (PlayerComputer(1).Checked = True) And (PlayerComputer(2).Checked = True) Then
			PlayerNames(1) = "Chip"
			PlayerNames(2) = "Meg"
		End If
		
		If (PlayerPerson(1).Checked = True) And (PlayerComputer(2).Checked = True) Then
			If PlayerNames(1) = "Chip" Then
				PlayerNames(2) = "Meg"
			End If
		End If
		
		If (PlayerPerson(2).Checked = True) And (PlayerComputer(1).Checked = True) Then
			If PlayerNames(2) = "Chip" Then
				PlayerNames(1) = "Meg"
			End If
		End If
		
		If PlayerNames(1) = PlayerNames(2) Then
			PlayerNames(1) = PlayerNames(1) & " (1)"
			PlayerNames(2) = PlayerNames(2) & " (2)"
		End If
		
	End Sub
	
	Private Sub CmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdOK.Click
		
		Call DefinePlayerNames()
		Call InitalOptions("Save")
		Hide()
		
	End Sub
	Private Sub CmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdCancel.Click
		
		Call InitalOptions("Restore")
		Call DefinePlayerNames()
		Hide()
		
	End Sub
End Class