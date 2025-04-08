VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'public class Tache

Option Explicit

Private id As Integer
Private intitule As String
Private duree As Integer
Private duree_nominale As Integer
Private ress As String
Private preds As String
Private fin As String
Private debut As String
Private t As Integer
Private debut_reel As Integer
Private colonne_fin_reelle As Integer

'Contructor
Private Sub Class_Initialize()
    id = taches.Count + 1
End Sub

'custom constructor
Public Sub set_attributes(i As String, d As Integer, r As String, p As String, Optional dpt As Integer)
    intitule = i
    duree_nominale = d
    ress = r 'doit pouvoir en avoir plusieurs
    preds = p
    If dpt = 0 Then
        duree = d / 2
    Else
        duree = dpt
    End If
End Sub

'affichage de la tâche sur la feuille excel
Public Sub Display()
    Dim s As Worksheet, h As Integer, v As Integer
    Set s = ThisWorkbook.Worksheets("TÂCHES")
    h = TSK_horizontal_margin
    v = TSK_vertical_margin
    Range(s.Cells(v - 1 + id, h + 1), s.Cells(v - 1 + id, h + 5)).Interior.Color = RGB(255, 255, 255)
    Range(s.Cells(v - 1 + id, h), s.Cells(v - 1 + id, h)).Interior.Color = RGB(220, 220, 220)
    Range(s.Cells(v - 1 + id, h), s.Cells(v - 1 + id, h + 5)).Borders.LineStyle = xlContinuous
    s.Cells(v - 1 + id, h) = id
    s.Cells(v - 1 + id, h + 1) = intitule
    s.Cells(v - 1 + id, h + 2) = duree_nominale / 8
    s.Cells(v - 1 + id, h + 3) = duree / 8
    s.Cells(v - 1 + id, h + 5) = ress
    s.Cells(v - 1 + id, h + 4) = preds
End Sub

'accesseurs
Public Property Get get_ID() As Integer
    get_ID = id
End Property
Public Property Get get_ress() As String
    get_ress = ress
End Property
Public Property Get get_preds() As String
    get_preds = preds
End Property
Public Property Get get_duree() As String
    get_duree = duree
End Property
Public Property Get get_debut_au_plus_tot() As String
    get_debut_au_plus_tot = debut_au_plus_tot
End Property
Public Property Get get_Intitule() As String
    get_Intitule = intitule
End Property
Public Property Get get_duree_nominale() As String
    get_duree_nominale = duree_nominale
End Property
Public Property Get get_debut() As String
    get_debut = debut
End Property
Public Property Get get_fin() As String
    get_fin = fin
End Property
Public Property Get get_type() As Integer
    get_type = t
End Property
Public Property Get get_colonne_fin_reelle() As String
    get_colonne_fin_reelle = colonne_fin_reelle
End Property
Public Property Get get_debut_reel() As String
    get_debut_reel = debut_reel
End Property

'setters
Public Sub set_type(value As Integer)
    t = value
End Sub

Public Sub set_ID(value As Integer)
    id = value
End Sub
Public Sub set_debut(value As Integer)
    debut = value
End Sub
Public Sub set_fin(value As Integer)
    fin = value
End Sub
Public Sub set_colonne_fin_reelle(value As Integer)
    colonne_fin_reelle = value
End Sub
Public Sub set_debut_reel(value As Integer)
    debut_reel = value
End Sub
Public Sub set_preds(value As String)
    preds = value
End Sub
Public Sub set_ress(value As String)
    ress = value
End Sub

'retourne une description de l'instance
Public Function str() As String
    str = "Tache numéro " + CStr(id) + ". Objet : " + intitule + "."
End Function
