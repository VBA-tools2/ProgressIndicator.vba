VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProgressIndicator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Const DEFAULT_CAPTION As String = "Progress"
Private Const DEFAULT_LABEL As String = "Please wait..."

Private Const ERR_NOT_INITIALIZED As String = "ProgressIndicator is not initialized."
Private Const ERR_PROC_NOT_FOUND As String = "Specified macro or object member was not found."
Private Const ERR_INVALID_OPERATION As String = "Worker procedure cannot be cancelled by assigning to this property."
Private Const VBERR_MEMBER_NOT_FOUND As Long = 438

Public Enum ProgressIndicatorError
    Error_NotInitialized = vbObjectError + 1001
    Error_ProcedureNotFound
    Error_InvalidOperation
End Enum

Private Type TProgressIndicator
    procedure As String
    instance As Object
    sleepDelay As Long
    canCancel As Boolean
    cancelling As Boolean
    currentProgressValue As Double
End Type

Private this As TProgressIndicator
Private WithEvents view As ProgressView
Attribute view.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set view = New ProgressView
    view.Caption = DEFAULT_CAPTION
    view.ProgressLabel = DEFAULT_LABEL
End Sub

Private Sub Class_Terminate()
    Set view = Nothing
    Set this.instance = Nothing
End Sub

Private Function QualifyMacroName(ByVal book As Workbook, ByVal procedure As String) As String
    QualifyMacroName = "'" & book.FullName & "'!" & procedure
End Function

Public Function Create(ByVal procedure As String, Optional instance As Object = Nothing, Optional ByVal initialLabelValue As String, Optional ByVal initialCaptionValue As String, Optional ByVal completedSleepMilliseconds As Long = 1000, Optional canCancel As Boolean = False) As ProgressIndicator
 
    Dim result As ProgressIndicator
    Set result = New ProgressIndicator
    
    result.Cancellable = canCancel
    result.SleepMilliseconds = completedSleepMilliseconds
 
    If Not instance Is Nothing Then
        Set result.OwnerInstance = instance
    ElseIf InStr(procedure, "'!") = 0 Then
        procedure = QualifyMacroName(Application.ActiveWorkbook, procedure)
    End If
 
    result.ProcedureName = procedure
 
    If initialLabelValue <> vbNullString Then result.ProgressView.ProgressLabel = initialLabelValue
    If initialCaptionValue <> vbNullString Then result.ProgressView.Caption = initialCaptionValue

    Set Create = result
 
End Function

Friend Property Get ProgressView() As ProgressView
    Set ProgressView = view
End Property

Friend Property Get ProcedureName() As String
    ProcedureName = this.procedure
End Property

Friend Property Let ProcedureName(ByVal value As String)
    this.procedure = value
End Property

Friend Property Get OwnerInstance() As Object
    Set OwnerInstance = this.instance
End Property

Friend Property Set OwnerInstance(ByVal value As Object)
    Set this.instance = value
End Property

Friend Property Get SleepMilliseconds() As Long
    SleepMilliseconds = this.sleepDelay
End Property

Friend Property Let SleepMilliseconds(ByVal value As Long)
    this.sleepDelay = value
End Property

Public Property Get CurrentProgress() As Double
    CurrentProgress = this.currentProgressValue
End Property

Public Property Get Cancellable() As Boolean
    Cancellable = this.canCancel
End Property

Friend Property Let Cancellable(ByVal value As Boolean)
    this.canCancel = value
End Property

Public Property Get IsCancelRequested() As Boolean
    IsCancelRequested = this.cancelling
End Property

Public Sub AbortCancellation()
    Debug.Assert this.cancelling
    this.cancelling = False
End Sub

Public Sub Execute()
    view.Show vbModal
End Sub

Public Sub Update(ByVal percentValue As Double, Optional ByVal labelValue As String, Optional ByVal captionValue As String)

    On Error GoTo CleanFail
    ThrowIfNotInitialized

    ValidatePercentValue percentValue
    this.currentProgressValue = percentValue
 
    view.Update this.currentProgressValue, labelValue

CleanExit:
    If percentValue = 1 Then Sleep 1000          ' pause on completion
    Exit Sub

CleanFail:
    MsgBox Err.Number & vbTab & Err.Description, vbCritical, "Error"
    Resume CleanExit
End Sub

Public Sub UpdatePercent(ByVal percentValue As Double, Optional ByVal captionValue As String)
    ValidatePercentValue percentValue
    Update percentValue, Format$(percentValue, "0.0% Completed")
End Sub

Private Sub ValidatePercentValue(ByRef percentValue As Double)
    If percentValue > 1 Then percentValue = percentValue / 100
End Sub

Private Sub ThrowIfNotInitialized()
    If this.procedure = vbNullString Then
        Err.Raise ProgressIndicatorError.Error_NotInitialized, TypeName(Me), ERR_NOT_INITIALIZED
    End If
End Sub

Private Sub view_Activated()

    On Error GoTo CleanFail
    ThrowIfNotInitialized

    If Not this.instance Is Nothing Then
        ExecuteInstanceMethod
    Else
        ExecuteMacro
    End If

CleanExit:
    view.Hide
    Exit Sub

CleanFail:
    MsgBox Err.Number & vbTab & Err.Description, vbCritical, "Error"
    Resume CleanExit
End Sub

Private Sub ExecuteMacro()
    On Error GoTo CleanFail
    Application.Run this.procedure, Me

CleanExit:
    Exit Sub

CleanFail:
    If Err.Number = VBERR_MEMBER_NOT_FOUND Then
        Err.Raise ProgressIndicatorError.Error_ProcedureNotFound, TypeName(Me), ERR_PROC_NOT_FOUND
    Else
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If
    Resume CleanExit
End Sub

Private Sub ExecuteInstanceMethod()
    On Error GoTo CleanFail
 
    Dim parameter As ProgressIndicator
    Set parameter = Me                           'Me cannot be passed to CallByName directly

    CallByName this.instance, this.procedure, VbMethod, parameter

CleanExit:
    Exit Sub

CleanFail:
    If Err.Number = VBERR_MEMBER_NOT_FOUND Then
        Err.Raise ProgressIndicatorError.Error_ProcedureNotFound, TypeName(Me), ERR_PROC_NOT_FOUND
    Else
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If
    Resume CleanExit
End Sub

Private Sub view_Cancelled()
    If Not this.canCancel Then Exit Sub
    this.cancelling = True
End Sub
