VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13560
   Icon            =   "Form1.frx":0000
   ScaleHeight     =   7770
   ScaleWidth      =   13560
   Begin VB.Frame Frame3 
      Caption         =   "Frame1"
      Height          =   6795
      Left            =   7620
      TabIndex        =   4
      Top             =   720
      Width           =   1875
      Begin MSComctlLib.ListView ListView3 
         Height          =   6375
         Left            =   120
         TabIndex        =   5
         Top             =   300
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   11245
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame1"
      Height          =   6795
      Left            =   4500
      TabIndex        =   2
      Top             =   720
      Width           =   1875
      Begin MSComctlLib.ListView ListView2 
         Height          =   6375
         Left            =   120
         TabIndex        =   3
         Top             =   300
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   11245
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   6795
      Left            =   1080
      TabIndex        =   0
      Top             =   540
      Width           =   1875
      Begin MSComctlLib.ListView ListView1 
         Height          =   6375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   11245
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Timer Timer_saida 
      Interval        =   5000
      Left            =   7080
      Top             =   840
   End
   Begin VB.Timer Timer_estacionado 
      Interval        =   10000
      Left            =   3900
      Top             =   780
   End
   Begin VB.Timer Timer_Entrada 
      Interval        =   1000
      Left            =   600
      Top             =   600
   End
   Begin VB.Label Label3 
      Caption         =   "Label1"
      Height          =   255
      Left            =   7800
      TabIndex        =   8
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Label1"
      Height          =   255
      Left            =   4680
      TabIndex        =   7
      Top             =   420
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim F_in As Collection
Dim F_stay As Collection
Dim F_out As Collection
Dim I As Integer
Dim qtde_vagas As Integer


Private Sub Form_Load()
Dim I As Integer
Dim X As ListItem

    qtde_vagas = 40
    
    Set F_out = New Collection
    Set F_stay = New Collection
    Set F_in = New Collection
    
    Me.ListView1.ListItems.Clear
    For I = 1 To 50
        F_in.Add "Carro " & I
        Set X = Me.ListView1.ListItems.Add(, , F_in(I))
    Next
End Sub

Private Sub Timer_Entrada_Timer()
Dim tempo1a5 As Integer
Dim X As ListItem


    'Transfere os carros da fila de entrada para a fila de estacionados
    tempo1a5 = (Rnd() * 5) + 1
    Timer_Entrada.Interval = 1000 * tempo1a5
    Label1.Caption = "Intervalo: " & Timer_Entrada.Interval
    
    If (F_in.Count > 0) And (F_stay.Count < qtde_vagas) Then
        F_stay.Add F_in(1)
        F_in.Remove (1)
        ListView2.ListItems.Clear
        For I = 1 To F_stay.Count
            Set X = Me.ListView2.ListItems.Add(, , F_stay(I))
        Next
    
        ListView1.ListItems.Clear
        For I = 1 To F_in.Count
            Set X = Me.ListView1.ListItems.Add(, , F_in(I))
        Next
        DoEvents
    End If
    
    
    tempo3a5 = (Rnd() * 2) + 3

    'Me.Label1 = Time() & vbCrLf & tempo1a5 & vbCrLf & tempo10a30 & vbCrLf & tempo3a5
End Sub

Private Sub Timer_estacionado_Timer()
Dim tempo10a30 As Integer
Dim X As ListItem

    'Transfere os carros estacionados para a fila de saída
    tempo10a30 = (Rnd() * 20) + 10
    Timer_estacionado.Interval = 1000 * tempo10a30
    Label2.Caption = "Intervalo: " & Timer_estacionado.Interval

    If F_stay.Count > 0 Then
        F_out.Add F_stay(1)
        F_stay.Remove (1)
        ListView3.ListItems.Clear
        For I = 1 To F_out.Count
            Set X = Me.ListView3.ListItems.Add(, , F_out(I))
        Next
    
        ListView2.ListItems.Clear
        For I = 1 To F_stay.Count
            Set X = Me.ListView2.ListItems.Add(, , F_stay(I))
        Next
        DoEvents
    End If

End Sub

Private Sub Timer_saida_Timer()
Dim tempo3a5 As Integer

    'transfere os carros da fila de saída para a fila de entrada
    tempo3a5 = (Rnd() * 2) + 3
    Timer_saida.Interval = 1000 * tempo3a5
    Label2.Caption = "Intervalo: " & Timer_saida.Interval

    If F_out.Count > 0 Then
        F_in.Add F_out(1)
        F_out.Remove (1)
        ListView1.ListItems.Clear
        For I = 1 To F_in.Count
            Set X = Me.ListView1.ListItems.Add(, , F_in(I))
        Next
    
        ListView3.ListItems.Clear
        For I = 1 To F_out.Count
            Set X = Me.ListView3.ListItems.Add(, , F_out(I))
        Next
        DoEvents
    End If

End Sub
