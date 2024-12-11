Option Strict Off
Option Explicit On
Friend Class frmAbout
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
	Public WithEvents CarewareTextBox As AxRichTextLib.AxRichTextBox
	Public WithEvents picIcon As System.Windows.Forms.PictureBox
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents LicenceTextBox As AxRichTextLib.AxRichTextBox
	Public WithEvents _Label1_1 As System.Windows.Forms.Label
	Public WithEvents _Label1_0 As System.Windows.Forms.Label
	Public WithEvents _Line1_2 As System.Windows.Forms.Label
	Public WithEvents _Line1_1 As System.Windows.Forms.Label
	Public WithEvents lblTitle As System.Windows.Forms.Label
	Public WithEvents lblDisclaimer As System.Windows.Forms.Label
	Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents Line1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmAbout))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ToolTip1.Active = True
		Me.CarewareTextBox = New AxRichTextLib.AxRichTextBox
		Me.picIcon = New System.Windows.Forms.PictureBox
		Me.cmdOK = New System.Windows.Forms.Button
		Me.LicenceTextBox = New AxRichTextLib.AxRichTextBox
		Me._Label1_1 = New System.Windows.Forms.Label
		Me._Label1_0 = New System.Windows.Forms.Label
		Me._Line1_2 = New System.Windows.Forms.Label
		Me._Line1_1 = New System.Windows.Forms.Label
		Me.lblTitle = New System.Windows.Forms.Label
		Me.lblDisclaimer = New System.Windows.Forms.Label
		Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.Line1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		CType(Me.CarewareTextBox, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.LicenceTextBox, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Line1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = " "
		Me.ClientSize = New System.Drawing.Size(430, 368)
		Me.Location = New System.Drawing.Point(156, 129)
		Me.Icon = CType(resources.GetObject("frmAbout.Icon"), System.Drawing.Icon)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmAbout"
		CarewareTextBox.OcxState = CType(resources.GetObject("CarewareTextBox.OcxState"), System.Windows.Forms.AxHost.State)
		Me.CarewareTextBox.Size = New System.Drawing.Size(402, 86)
		Me.CarewareTextBox.Location = New System.Drawing.Point(15, 81)
		Me.CarewareTextBox.TabIndex = 4
		Me.CarewareTextBox.Name = "CarewareTextBox"
		Me.picIcon.Size = New System.Drawing.Size(36, 36)
		Me.picIcon.Location = New System.Drawing.Point(14, 16)
		Me.picIcon.Image = CType(resources.GetObject("picIcon.Image"), System.Drawing.Image)
		Me.picIcon.TabIndex = 1
		Me.picIcon.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picIcon.Dock = System.Windows.Forms.DockStyle.None
		Me.picIcon.BackColor = System.Drawing.SystemColors.Control
		Me.picIcon.CausesValidation = True
		Me.picIcon.Enabled = True
		Me.picIcon.ForeColor = System.Drawing.SystemColors.ControlText
		Me.picIcon.Cursor = System.Windows.Forms.Cursors.Default
		Me.picIcon.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picIcon.TabStop = True
		Me.picIcon.Visible = True
		Me.picIcon.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
		Me.picIcon.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.picIcon.Name = "picIcon"
		Me.cmdOK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.cmdOK
		Me.cmdOK.Text = "OK"
		Me.AcceptButton = Me.cmdOK
		Me.cmdOK.Size = New System.Drawing.Size(84, 23)
		Me.cmdOK.Location = New System.Drawing.Point(320, 333)
		Me.cmdOK.TabIndex = 0
		Me.cmdOK.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOK.CausesValidation = True
		Me.cmdOK.Enabled = True
		Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOK.TabStop = True
		Me.cmdOK.Name = "cmdOK"
		LicenceTextBox.OcxState = CType(resources.GetObject("LicenceTextBox.OcxState"), System.Windows.Forms.AxHost.State)
		Me.LicenceTextBox.Size = New System.Drawing.Size(402, 108)
		Me.LicenceTextBox.Location = New System.Drawing.Point(14, 197)
		Me.LicenceTextBox.TabIndex = 5
		Me.LicenceTextBox.Name = "LicenceTextBox"
		Me._Label1_1.Text = "Licence:"
		Me._Label1_1.Size = New System.Drawing.Size(401, 15)
		Me._Label1_1.Location = New System.Drawing.Point(14, 178)
		Me._Label1_1.TabIndex = 7
		Me._Label1_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_1.BackColor = System.Drawing.SystemColors.Control
		Me._Label1_1.Enabled = True
		Me._Label1_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_1.UseMnemonic = True
		Me._Label1_1.Visible = True
		Me._Label1_1.AutoSize = False
		Me._Label1_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_1.Name = "_Label1_1"
		Me._Label1_0.Text = "Careware:"
		Me._Label1_0.Size = New System.Drawing.Size(401, 15)
		Me._Label1_0.Location = New System.Drawing.Point(14, 63)
		Me._Label1_0.TabIndex = 6
		Me._Label1_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_0.BackColor = System.Drawing.SystemColors.Control
		Me._Label1_0.Enabled = True
		Me._Label1_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_0.UseMnemonic = True
		Me._Label1_0.Visible = True
		Me._Label1_0.AutoSize = False
		Me._Label1_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_0.Name = "_Label1_0"
		Me._Line1_2.BackColor = System.Drawing.Color.FromARGB(128, 128, 128)
		Me._Line1_2.Visible = True
		Me._Line1_2.Location = New System.Drawing.Point(11, 33)
		Me._Line1_2.Width = 291
		Me._Line1_2.Height = 1
		Me._Line1_2.Name = "_Line1_2"
		Me._Line1_1.BackColor = System.Drawing.Color.FromARGB(128, 128, 128)
		Me._Line1_1.Visible = True
		Me._Line1_1.Location = New System.Drawing.Point(13, 185)
		Me._Line1_1.Width = 288
		Me._Line1_1.Height = 1
		Me._Line1_1.Name = "_Line1_1"
		Me.lblTitle.Text = "Concentration"
		Me.lblTitle.ForeColor = System.Drawing.Color.Black
		Me.lblTitle.Size = New System.Drawing.Size(298, 20)
		Me.lblTitle.Location = New System.Drawing.Point(65, 25)
		Me.lblTitle.TabIndex = 3
		Me.lblTitle.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblTitle.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblTitle.BackColor = System.Drawing.SystemColors.Control
		Me.lblTitle.Enabled = True
		Me.lblTitle.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblTitle.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblTitle.UseMnemonic = True
		Me.lblTitle.Visible = True
		Me.lblTitle.AutoSize = False
		Me.lblTitle.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblTitle.Name = "lblTitle"
		Me.lblDisclaimer.Text = "Copyright 1998 - 2002 Rob Latour"
		Me.lblDisclaimer.ForeColor = System.Drawing.Color.Black
		Me.lblDisclaimer.Size = New System.Drawing.Size(258, 19)
		Me.lblDisclaimer.Location = New System.Drawing.Point(14, 336)
		Me.lblDisclaimer.TabIndex = 2
		Me.lblDisclaimer.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblDisclaimer.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblDisclaimer.BackColor = System.Drawing.SystemColors.Control
		Me.lblDisclaimer.Enabled = True
		Me.lblDisclaimer.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblDisclaimer.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblDisclaimer.UseMnemonic = True
		Me.lblDisclaimer.Visible = True
		Me.lblDisclaimer.AutoSize = False
		Me.lblDisclaimer.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblDisclaimer.Name = "lblDisclaimer"
		Me.Controls.Add(CarewareTextBox)
		Me.Controls.Add(picIcon)
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(LicenceTextBox)
		Me.Controls.Add(_Label1_1)
		Me.Controls.Add(_Label1_0)
		Me.Controls.Add(_Line1_2)
		Me.Controls.Add(_Line1_1)
		Me.Controls.Add(lblTitle)
		Me.Controls.Add(lblDisclaimer)
		Me.Label1.SetIndex(_Label1_1, CType(1, Short))
		Me.Label1.SetIndex(_Label1_0, CType(0, Short))
		Me.Line1.SetIndex(_Line1_2, CType(2, Short))
		Me.Line1.SetIndex(_Line1_1, CType(1, Short))
		CType(Me.Line1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.LicenceTextBox, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.CarewareTextBox, System.ComponentModel.ISupportInitialize).EndInit()
	End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As frmAbout
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As frmAbout
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New frmAbout()
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
	
	Private Sub CmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdOK.Click
		
		Me.Close()
		
	End Sub
	
	'UPGRADE_WARNING: Form event frmAbout.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2065"'
	Private Sub frmAbout_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		
		Const NewParagraph As String = vbLf & vbLf
		
		Me.Text = System.Reflection.Assembly.GetExecutingAssembly.GetName.Name & " - About"
		
		'UPGRADE_ISSUE: App property App.Revision was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2069"'
        'lblTitle.Text = System.Reflection.Assembly.GetExecutingAssembly.GetName.Name & " Version " & System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location).FileMajorPart & "." & System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location).FileMinorPart & "." & App.Revision & " (32bit)"
		
		CarewareTextBox.Text = "I'm sure you've heard of the shareware concept, where the author of " & "a program for his or her hard work requests that you send them a payment.  " & "With this program I would like to introduce you to the concept of Careware.  " & "If you plan to keep and/or use this program, Concentration, then you are " & "asked to donate $5 to someone much less fortunate than yourself."
		
		LicenceTextBox.Text = "This Concentration Software (the 'SOFTWARE') and documentation are owned by Rob Latour. Rob Latour grants to you a nonexclusive license to use the SOFTWARE and documentation solely for your personal, family, and internal business purposes. You may not commercially distribute, sublicense, resell, or otherwise transfer for any consideration, or reproduce for any such purposes, the SOFTWARE and documentation or any modification or derivation thereof, either alone or in conjunction with any other product or program without the prior written authorization of Rob Latour. Further, you may not modify or translate the SOFTWARE or documentation, other than for your own use. " & NewParagraph & "The SOFTWARE and documentation are provided 'as is'. Rob Latour makes no warranties, either express or implied, with respect to the SOFTWARE, documentation, and associated materials provided to you, including but not limited to any warranty of merchantability, fitness for a particular purpose or against infringement. Rob Latour does not warrant that the functions contained in the SOFTWARE will meet your requirements, or that the operation of the SOFTWARE will be uninterrupted or error-free, or that defects in the SOFTWARE or documentation will be corrected. Furthermore, Rob Latour does not warrant or make any representations regarding the use or the results of the use of the SOFTWARE or documentation provided therewith in terms of their correctness, accuracy, reliability, or otherwise. No oral or written information or advice given by Rob Latour or authorized representative of Rob Latour shall create a warranty or in any way increase the scope of this warranty." & NewParagraph & "Limitation of liability - Rob Latour and his licensors are not liable for any claims or damages whatsoever, including property damage, personal injury, intellectual property infringement, loss of profits, or interruption of business, or for any special, consequential or incidental damages, however caused, whether arising out of breach of warranty, contract, tort (including negligence), strict liability, or otherwise." & NewParagraph & "The SOFTWARE and documentation was written by Rob Latour; the SOFTWARE and documentation are Copyright 1998 - 2002 by Rob Latour." & NewParagraph & "© 1998 - 2002 Rob Latour.  All rights reserved." & NewParagraph & "This SOFTWARE employs various Windows components, including CARDS32.DLL, copyrighted by Microsoft. Microsoft components not already licensed to standard Windows licensees, used by this Software, have been licensed for redistribution with this SOFTWARE. Windows is a registered trademark of Microsoft Corporation. © 1998 Microsoft Corporation.  All rights reserved." & NewParagraph & "All other trademarks and service marks are the property of their respective owners."
		
		
	End Sub
End Class