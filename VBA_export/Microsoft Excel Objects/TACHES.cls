VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tâches"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit
Public up As Boolean 'si Faux, l'actualisation des ressources en cas de modification de la feuille est désactivée

Private Sub Worksheet_Activate()
    up = True
    If Cells(2, 1).value = "" Then
        MsgBox "Aucune date de début renseignée. Veuillez indiquer la date de lancement du projet en cellule A2!"
    End If
    
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    
    'ctr = ctr + 1
    'Set test = New Collection
    'MsgBox test.Count
    
    Dim r As Range
    Set r = Range(Cells(2, 1), Cells(2, 1))
    If Not Intersect(Target, r) Is Nothing Then
        With Worksheets("GANTT")
            If .Cells(2, 3).value <> Target.value Then
                .Cells(2, 3) = Target.value
            End If
        End With
        With Worksheets("DASHBOARD")
        If .Cells(2, 3).value <> Target.value Then
            .Cells(2, 3) = Target.value
        End If
        End With
    End If
    If up = True Then
        Dim r2 As Range, w As Integer: w = 10
        While Cells(w, 7).Interior.Color = vbWhite
            w = w + 1
        Wend
        Set r2 = Range(Cells(10, 7), Cells(w - 1, 7))
        If Not Intersect(Target, r2) Is Nothing Then
            Call actualiser
        End If
    End If
End Sub

