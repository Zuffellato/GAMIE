
Partial Public Class FRMZCIMAN
    Inherits FRM__CHIL

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New()
    End Sub

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub
  Friend WithEvents GridColumn1_981_6 As NTSGridColumn
  Friend WithEvents GridColumn1_406_7 As NTSGridColumn
  Friend WithEvents GridColumn1_824_8 As NTSGridColumn
  Friend WithEvents GridColumn1_382_9 As NTSGridColumn
  Friend WithEvents GridColumn1_240_10 As NTSGridColumn
  Friend WithEvents GridColumn1_296_11 As NTSGridColumn
  Friend WithEvents GridColumn1_776_12 As NTSGridColumn
  Friend WithEvents GridColumn1_895_13 As NTSGridColumn
  Friend WithEvents GridColumn1_767_14 As NTSGridColumn
  Friend WithEvents tlbApri As NTSBarButtonItem
  Friend WithEvents tlbCancMultipla As NTSBarButtonItem
  Friend WithEvents tlbDuplica As NTSBarButtonItem
    Public WithEvents tlbElabora As NTSBarButtonItem
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents ck_maeriali As NTSCheckBox
    Friend WithEvents ck_mezzi As NTSCheckBox
    Friend WithEvents ck_risorse As NTSCheckBox
End Class