Attribute VB_Name = "modWorkerCode"

Option Explicit

Public Sub DoSomething()
    With ProgressIndicator.Create("DoWork", canCancel:=True)
        .Execute
    End With
End Sub

Public Sub DoWork(ByVal progress As ProgressIndicator)
    Dim i As Long
    For i = 1 To 10000
        If ShouldCancel(progress) Then
            'here more complex worker code could rollback & cleanup
            Exit Sub
        End If
        ActiveSheet.Cells(1, 1) = i
        progress.Update i / 10000              'show only the bar
'        progress.UpdatePercent i / 10000       'show also percentage value
    Next
End Sub

Private Function ShouldCancel(ByVal progress As ProgressIndicator) As Boolean
    If progress.IsCancelRequested Then
        If MsgBox("Cancel this operation?", vbYesNo) = vbYes Then
            ShouldCancel = True
        Else
            progress.AbortCancellation
        End If
    End If
End Function
