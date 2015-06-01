VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainFrm 
   Caption         =   "Trabalho de SI 2015/1"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9840.001
   OleObjectBlob   =   "MainFrm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'MainFrm file
'Written by João Marcelo Brito
'10/05/2015
'=====================================================================

Private Sub btnTimer_Click()
    Dim i As Integer
    
    Worksheets("MainSheet").Cells(1, 2) = cbStock.Text
    If (btnTimer.Value = True) Then
        If (txtDelay.Text = vbNullString Or cbStock.Text = vbNullString) Then
            MsgBox "Por favor digite o delay corretamente e o código da ação"
        Else
            btnTimer.Caption = "Desativar Timer"
            bTimer = True
            StockTimer
        End If
    Else
        btnTimer.Caption = "Ativar Timer"
        bTimer = False
    End If
End Sub

Private Sub cbGraph_Click()
    
    If cbGraph.Value = True Then
        Worksheets("MainSheet").ChartObjects(1).Visible = True
    Else
        Worksheets("MainSheet").ChartObjects(1).Visible = False
    End If
End Sub

Private Sub txtDelay_Change()
    If Not IsNumeric(txtDelay.Value) And txtDelay.Value <> vbNullString Then
        MsgBox "Este campo só aceita números."
        txtDelay.Text = vbNullString
    End If
End Sub
'End of file
'=====================================================================================================================
