VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ressource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private id As Integer
Private nom As String
Private t As String

'Contructor
Private Sub Class_Initialize()
    id = ressources.Count + 1
End Sub

'custom constructor
Public Sub set_attributes(n As String, Optional tach As String)
    nom = n
    If Not IsMissing(tach) Then
        t = tach
    End If
End Sub

'affichage de la tâche sur la feuille excel
Public Sub Display()
    Dim s As Worksheet, v As Integer, h As Integer
    Set s = ThisWorkbook.Worksheets("TÂCHES")
    v = RSC_vertical_margin
    h = RSC_horizontal_margin
    Range(s.Cells(v - 1 + id, h), s.Cells(v - 1 + id, h)).Interior.Color = RGB(220, 220, 220)
    Range(s.Cells(v - 1 + id, h + 1), s.Cells(v - 1 + id, h + 1)).Interior.Color = RGB(255, 255, 255)
    Range(s.Cells(v - 1 + id, h + 2), s.Cells(v - 1 + id, h + 2)).Interior.Color = RGB(220, 220, 220)
    Range(s.Cells(v - 1 + id, h), s.Cells(v - 1 + id, h + 2)).Borders.LineStyle = xlContinuous
    s.Cells(v - 1 + id, h) = Chr(64 + id)
    s.Cells(v - 1 + id, h + 1) = nom
End Sub

'accesseurs
Public Property Get get_ID() As Integer
    get_ID = id
End Property


Public Property Get get_Name() As String
    get_Name = nom
End Property
Public Property Get get_t() As String
    get_t = t
End Property

Public Sub set_ID(value As Integer)
    id = value
End Sub
Public Sub set_t(value As String)
    t = value
End Sub


'retourne une description de l'instance
Public Function str() As String
    str = "Ressource " + Chr(64 + id) + ". Nom : " + nom + "."
End Function

