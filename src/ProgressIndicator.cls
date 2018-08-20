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

'==============================================================================
Private Const DEFAULT_CAPTION As String = "Progress"
Private Const DEFAULT_LABEL As String = "Please wait..."

Private Const ERR_NOT_INITIALIZED As String = _
        "ProgressIndicator is not initialized."
Private Const ERR_PROC_NOT_FOUND As String = _
        "Specified macro or object member was not found."
Private Const ERR_INVALID_OPERATION As String = _
        "Worker procedure cannot be canceled by assigning to this property."
Private Const VBERR_MEMBER_NOT_FOUND As Long = 438
'==============================================================================

Public Enum ProgressIndicatorError
    Error_NotInitialized = vbObjectError + 1001
    Error_ProcedureNotFound
End Enum

Private Type TProgressIndicator
    procedure As String
    instance As Object
    sleepDelay As Long
    canCancel As Boolean
    cancelling As Boolean
    currentProgressValue As Double
End Type

Private This As TProgressIndicator
Private WithEvents View As ProgressView
Attribute View.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set View = New ProgressView
    View.Caption = DEFAULT_CAPTION
    View.ProgressLabel = DEFAULT_LABEL
End Sub

Private Sub Class_Terminate()
    Set View = Nothing
    Set This.instance = Nothing
End Sub

Private Function QualifyMacroName( _
    ByVal book As Workbook, _
    ByVal procedure As String _
        ) As String
    QualifyMacroName = "'" & book.FullName & "'!" & procedure
End Function

Public Function Create( _
    ByVal procedure As String, _
    Optional instance As Object = Nothing, _
    Optional ByVal initialLabelValue As String, _
    Optional ByVal initialCaptionValue As String, _
    Optional ByVal completedSleepMilliseconds As Long = 1000, _
    Optional canCancel As Boolean = False _
        ) As ProgressIndicator
    
    Dim Result As ProgressIndicator
    Set Result = New ProgressIndicator
    
    Result.Cancellable = canCancel
    Result.SleepMilliseconds = completedSleepMilliseconds
    
    If Not instance Is Nothing Then
        Set Result.OwnerInstance = instance
    ElseIf InStr(procedure, "'!") = 0 Then
        procedure = QualifyMacroName(Application.ActiveWorkbook, procedure)
    End If
    
    Result.ProcedureName = procedure
    
    If initialLabelValue <> vbNullString Then _
            Result.ProgressView.ProgressLabel = initialLabelValue
    If initialCaptionValue <> vbNullString Then _
            Result.ProgressView.Caption = initialCaptionValue
    
    Set Create = Result
    
End Function

Friend Property Get ProgressView() As ProgressView
    Set ProgressView = View
End Property

Friend Property Get ProcedureName() As String
    ProcedureName = This.procedure
End Property

Friend Property Let ProcedureName(ByVal Value As String)
    This.procedure = Value
End Property

Friend Property Get OwnerInstance() As Object
    Set OwnerInstance = This.instance
End Property

Friend Property Set OwnerInstance(ByVal Value As Object)
    Set This.instance = Value
End Property

Friend Property Get SleepMilliseconds() As Long
    SleepMilliseconds = This.sleepDelay
End Property

Friend Property Let SleepMilliseconds(ByVal Value As Long)
    This.sleepDelay = Value
End Property

Public Property Get CurrentProgress() As Double
    CurrentProgress = This.currentProgressValue
End Property

Public Property Get Cancellable() As Boolean
    Cancellable = This.canCancel
End Property

Friend Property Let Cancellable(ByVal Value As Boolean)
    This.canCancel = Value
End Property

Public Property Get IsCancelRequested() As Boolean
    IsCancelRequested = This.cancelling
End Property

Public Sub AbortCancellation()
Debug.Assert This.cancelling
    This.cancelling = False
End Sub

Public Sub Execute()
    View.Show vbModal
End Sub

Public Sub Update( _
    ByVal percentValue As Double, _
    Optional ByVal labelValue As String, _
    Optional ByVal captionValue As String _
)
    
    On Error GoTo CleanFail
    ThrowIfNotInitialized
    
    ValidatePercentValue percentValue
    This.currentProgressValue = percentValue
    
    View.Update This.currentProgressValue, labelValue
    
CleanExit:
    If percentValue = 1 Then Sleep 1000          'pause on completion
    Exit Sub
    
CleanFail:
    MsgBox Err.Number & vbTab & Err.Description, vbCritical, "Error"
    Resume CleanExit
End Sub

Public Sub UpdatePercent( _
    ByVal percentValue As Double, _
    Optional ByVal captionValue As String _
)
    ValidatePercentValue percentValue
    Update percentValue, Format$(percentValue, "0.0% Completed")
End Sub

Private Sub ValidatePercentValue(ByRef percentValue As Double)
    If percentValue > 1 Then percentValue = percentValue / 100
End Sub

Private Sub ThrowIfNotInitialized()
    If This.procedure = vbNullString Then
        Err.Raise ProgressIndicatorError.Error_NotInitialized, _
                TypeName(Me), ERR_NOT_INITIALIZED
    End If
End Sub

Private Sub View_Activated()
    
    On Error GoTo CleanFail
    ThrowIfNotInitialized
    
    If Not This.instance Is Nothing Then
        ExecuteInstanceMethod
    Else
        ExecuteMacro
    End If
    
CleanExit:
    View.Hide
    Exit Sub
    
CleanFail:
    MsgBox Err.Number & vbTab & Err.Description, vbCritical, "Error"
    Resume CleanExit
End Sub

Private Sub ExecuteMacro()
    On Error GoTo CleanFail
    Application.Run This.procedure, Me
    
CleanExit:
    Exit Sub
    
CleanFail:
    If Err.Number = VBERR_MEMBER_NOT_FOUND Then
        Err.Raise ProgressIndicatorError.Error_ProcedureNotFound, _
                TypeName(Me), ERR_PROC_NOT_FOUND
    Else
        Err.Raise _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                Err.HelpFile, _
                Err.HelpContext
    End If
    Resume CleanExit
End Sub

Private Sub ExecuteInstanceMethod()
    On Error GoTo CleanFail
    
    Dim parameter As ProgressIndicator
    Set parameter = Me            'Me cannot be passed to CallByName directly
    
    CallByName This.instance, This.procedure, VbMethod, parameter
    
CleanExit:
    Exit Sub
    
CleanFail:
    If Err.Number = VBERR_MEMBER_NOT_FOUND Then
        Err.Raise _
                ProgressIndicatorError.Error_ProcedureNotFound, _
                TypeName(Me), _
                ERR_PROC_NOT_FOUND
    Else
        Err.Raise _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                Err.HelpFile, _
                Err.HelpContext
    End If
    Resume CleanExit
End Sub

Private Sub View_Cancelled()
    If Not This.canCancel Then Exit Sub
    This.cancelling = True
End Sub
