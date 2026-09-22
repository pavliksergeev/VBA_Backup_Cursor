VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSoilAvg
   Caption         =   "SoilAvg"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9900
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSoilAvg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' Mac paste: Insert UserForm, name frmSoilAvg, paste from Option Explicit below.
' Windows: File -> Import this .frm, then Module_SoilAvg.bas.
'
Option Explicit

Private WithEvents btnAdd As MSForms.CommandButton
Private WithEvents btnRemove As MSForms.CommandButton
Private WithEvents btnCalc As MSForms.CommandButton
Private WithEvents btnApply As MSForms.CommandButton
Private WithEvents btnSave As MSForms.CommandButton
Private WithEvents btnClose As MSForms.CommandButton

Private Sub UserForm_Initialize()
    Call SoilAvg_BuildUi(Me)
    On Error Resume Next
    Set btnAdd = Me.Controls("btnAdd")
    Set btnRemove = Me.Controls("btnRemove")
    Set btnCalc = Me.Controls("btnCalc")
    Set btnApply = Me.Controls("btnApply")
    Set btnSave = Me.Controls("btnSave")
    Set btnClose = Me.Controls("btnClose")
    On Error GoTo 0
End Sub

Private Sub btnAdd_Click()
    Call SoilAvg_OnAdd
End Sub

Private Sub btnRemove_Click()
    Call SoilAvg_OnRemove
End Sub

Private Sub btnCalc_Click()
    Call SoilAvg_OnCalc
End Sub

Private Sub btnApply_Click()
    Call SoilAvg_OnApply
End Sub

Private Sub btnSave_Click()
    Call SoilAvg_OnSave
End Sub

Private Sub btnClose_Click()
    Call SoilAvg_OnClose
End Sub
