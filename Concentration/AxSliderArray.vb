'UPGRADE_WARNING: The entire project must be compiled once before a form with an ActiveX Control Array can be displayed

Imports System.ComponentModel

<ProvideProperty("Index",GetType(AxComctlLib.AxSlider))> Public Class AxSliderArray
	Inherits Microsoft.VisualBasic.Compatibility.VB6.BaseOcxArray
	Implements IExtenderProvider

	Public Sub New()
		MyBase.New()
	End Sub

	Public Sub New(ByVal Container As IContainer)
		MyBase.New(Container)
	End Sub

	Public Shadows Event [ClickEvent] (ByVal sender As System.Object, ByVal e As System.EventArgs)
	Public Shadows Event [KeyDownEvent] (ByVal sender As System.Object, ByVal e As AxComctlLib.ISliderEvents_KeyDownEvent)
	Public Shadows Event [KeyPressEvent] (ByVal sender As System.Object, ByVal e As AxComctlLib.ISliderEvents_KeyPressEvent)
	Public Shadows Event [KeyUpEvent] (ByVal sender As System.Object, ByVal e As AxComctlLib.ISliderEvents_KeyUpEvent)
	Public Shadows Event [MouseDownEvent] (ByVal sender As System.Object, ByVal e As AxComctlLib.ISliderEvents_MouseDownEvent)
	Public Shadows Event [MouseMoveEvent] (ByVal sender As System.Object, ByVal e As AxComctlLib.ISliderEvents_MouseMoveEvent)
	Public Shadows Event [MouseUpEvent] (ByVal sender As System.Object, ByVal e As AxComctlLib.ISliderEvents_MouseUpEvent)
	Public Shadows Event [Scroll] (ByVal sender As System.Object, ByVal e As System.EventArgs)
	Public Shadows Event [Change] (ByVal sender As System.Object, ByVal e As System.EventArgs)
	Public Shadows Event [OLEStartDrag] (ByVal sender As System.Object, ByVal e As AxComctlLib.ISliderEvents_OLEStartDragEvent)
	Public Shadows Event [OLEGiveFeedback] (ByVal sender As System.Object, ByVal e As AxComctlLib.ISliderEvents_OLEGiveFeedbackEvent)
	Public Shadows Event [OLESetData] (ByVal sender As System.Object, ByVal e As AxComctlLib.ISliderEvents_OLESetDataEvent)
	Public Shadows Event [OLECompleteDrag] (ByVal sender As System.Object, ByVal e As AxComctlLib.ISliderEvents_OLECompleteDragEvent)
	Public Shadows Event [OLEDragOver] (ByVal sender As System.Object, ByVal e As AxComctlLib.ISliderEvents_OLEDragOverEvent)
	Public Shadows Event [OLEDragDrop] (ByVal sender As System.Object, ByVal e As AxComctlLib.ISliderEvents_OLEDragDropEvent)

	<System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)> Public Function CanExtend(ByVal target As Object) As Boolean Implements IExtenderProvider.CanExtend
		If TypeOf target Is AxComctlLib.AxSlider Then
			Return BaseCanExtend(target)
		End If
	End Function

	Public Function GetIndex(ByVal o As AxComctlLib.AxSlider) As Short
		Return BaseGetIndex(o)
	End Function

	<System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)> Public Sub SetIndex(ByVal o As AxComctlLib.AxSlider, ByVal Index As Short)
		BaseSetIndex(o, Index)
	End Sub

	<System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)> Public Function ShouldSerializeIndex(ByVal o As AxComctlLib.AxSlider) As Boolean
		Return BaseShouldSerializeIndex(o)
	End Function

	<System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)> Public Sub ResetIndex(ByVal o As AxComctlLib.AxSlider)
		BaseResetIndex(o)
	End Sub

	Default Public ReadOnly Property Item(ByVal Index As Short) As AxComctlLib.AxSlider
		Get
			Item = CType(BaseGetItem(Index), AxComctlLib.AxSlider)
		End Get
	End Property

	Protected Overrides Function GetControlInstanceType() As System.Type
		Return GetType(AxComctlLib.AxSlider)
	End Function

	Protected Overrides Sub HookUpControlEvents(ByVal o As Object)
		Dim ctl As AxComctlLib.AxSlider = CType(o, AxComctlLib.AxSlider)
		MyBase.HookUpControlEvents(o)
		If Not ClickEventEvent Is Nothing Then
			AddHandler ctl.ClickEvent, New System.EventHandler(AddressOf HandleClickEvent)
		End If
		If Not KeyDownEventEvent Is Nothing Then
			AddHandler ctl.KeyDownEvent, New AxComctlLib.ISliderEvents_KeyDownEventHandler(AddressOf HandleKeyDownEvent)
		End If
		If Not KeyPressEventEvent Is Nothing Then
			AddHandler ctl.KeyPressEvent, New AxComctlLib.ISliderEvents_KeyPressEventHandler(AddressOf HandleKeyPressEvent)
		End If
		If Not KeyUpEventEvent Is Nothing Then
			AddHandler ctl.KeyUpEvent, New AxComctlLib.ISliderEvents_KeyUpEventHandler(AddressOf HandleKeyUpEvent)
		End If
		If Not MouseDownEventEvent Is Nothing Then
			AddHandler ctl.MouseDownEvent, New AxComctlLib.ISliderEvents_MouseDownEventHandler(AddressOf HandleMouseDownEvent)
		End If
		If Not MouseMoveEventEvent Is Nothing Then
			AddHandler ctl.MouseMoveEvent, New AxComctlLib.ISliderEvents_MouseMoveEventHandler(AddressOf HandleMouseMoveEvent)
		End If
		If Not MouseUpEventEvent Is Nothing Then
			AddHandler ctl.MouseUpEvent, New AxComctlLib.ISliderEvents_MouseUpEventHandler(AddressOf HandleMouseUpEvent)
		End If
		If Not ScrollEvent Is Nothing Then
			AddHandler ctl.Scroll, New System.EventHandler(AddressOf HandleScroll)
		End If
		If Not ChangeEvent Is Nothing Then
			AddHandler ctl.Change, New System.EventHandler(AddressOf HandleChange)
		End If
		If Not OLEStartDragEvent Is Nothing Then
			AddHandler ctl.OLEStartDrag, New AxComctlLib.ISliderEvents_OLEStartDragEventHandler(AddressOf HandleOLEStartDrag)
		End If
		If Not OLEGiveFeedbackEvent Is Nothing Then
			AddHandler ctl.OLEGiveFeedback, New AxComctlLib.ISliderEvents_OLEGiveFeedbackEventHandler(AddressOf HandleOLEGiveFeedback)
		End If
		If Not OLESetDataEvent Is Nothing Then
			AddHandler ctl.OLESetData, New AxComctlLib.ISliderEvents_OLESetDataEventHandler(AddressOf HandleOLESetData)
		End If
		If Not OLECompleteDragEvent Is Nothing Then
			AddHandler ctl.OLECompleteDrag, New AxComctlLib.ISliderEvents_OLECompleteDragEventHandler(AddressOf HandleOLECompleteDrag)
		End If
		If Not OLEDragOverEvent Is Nothing Then
			AddHandler ctl.OLEDragOver, New AxComctlLib.ISliderEvents_OLEDragOverEventHandler(AddressOf HandleOLEDragOver)
		End If
		If Not OLEDragDropEvent Is Nothing Then
			AddHandler ctl.OLEDragDrop, New AxComctlLib.ISliderEvents_OLEDragDropEventHandler(AddressOf HandleOLEDragDrop)
		End If
	End Sub

	Private Sub HandleClickEvent (ByVal sender As System.Object, ByVal e As System.EventArgs) 
		RaiseEvent [ClickEvent] (sender, e)
	End Sub

	Private Sub HandleKeyDownEvent (ByVal sender As System.Object, ByVal e As AxComctlLib.ISliderEvents_KeyDownEvent) 
		RaiseEvent [KeyDownEvent] (sender, e)
	End Sub

	Private Sub HandleKeyPressEvent (ByVal sender As System.Object, ByVal e As AxComctlLib.ISliderEvents_KeyPressEvent) 
		RaiseEvent [KeyPressEvent] (sender, e)
	End Sub

	Private Sub HandleKeyUpEvent (ByVal sender As System.Object, ByVal e As AxComctlLib.ISliderEvents_KeyUpEvent) 
		RaiseEvent [KeyUpEvent] (sender, e)
	End Sub

	Private Sub HandleMouseDownEvent (ByVal sender As System.Object, ByVal e As AxComctlLib.ISliderEvents_MouseDownEvent) 
		RaiseEvent [MouseDownEvent] (sender, e)
	End Sub

	Private Sub HandleMouseMoveEvent (ByVal sender As System.Object, ByVal e As AxComctlLib.ISliderEvents_MouseMoveEvent) 
		RaiseEvent [MouseMoveEvent] (sender, e)
	End Sub

	Private Sub HandleMouseUpEvent (ByVal sender As System.Object, ByVal e As AxComctlLib.ISliderEvents_MouseUpEvent) 
		RaiseEvent [MouseUpEvent] (sender, e)
	End Sub

	Private Sub HandleScroll (ByVal sender As System.Object, ByVal e As System.EventArgs) 
		RaiseEvent [Scroll] (sender, e)
	End Sub

	Private Sub HandleChange (ByVal sender As System.Object, ByVal e As System.EventArgs) 
		RaiseEvent [Change] (sender, e)
	End Sub

	Private Sub HandleOLEStartDrag (ByVal sender As System.Object, ByVal e As AxComctlLib.ISliderEvents_OLEStartDragEvent) 
		RaiseEvent [OLEStartDrag] (sender, e)
	End Sub

	Private Sub HandleOLEGiveFeedback (ByVal sender As System.Object, ByVal e As AxComctlLib.ISliderEvents_OLEGiveFeedbackEvent) 
		RaiseEvent [OLEGiveFeedback] (sender, e)
	End Sub

	Private Sub HandleOLESetData (ByVal sender As System.Object, ByVal e As AxComctlLib.ISliderEvents_OLESetDataEvent) 
		RaiseEvent [OLESetData] (sender, e)
	End Sub

	Private Sub HandleOLECompleteDrag (ByVal sender As System.Object, ByVal e As AxComctlLib.ISliderEvents_OLECompleteDragEvent) 
		RaiseEvent [OLECompleteDrag] (sender, e)
	End Sub

	Private Sub HandleOLEDragOver (ByVal sender As System.Object, ByVal e As AxComctlLib.ISliderEvents_OLEDragOverEvent) 
		RaiseEvent [OLEDragOver] (sender, e)
	End Sub

	Private Sub HandleOLEDragDrop (ByVal sender As System.Object, ByVal e As AxComctlLib.ISliderEvents_OLEDragDropEvent) 
		RaiseEvent [OLEDragDrop] (sender, e)
	End Sub

End Class

