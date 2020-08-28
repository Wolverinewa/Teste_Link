VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10425
   Icon            =   "Form1.frx":0000
   ScaleHeight     =   7665
   ScaleWidth      =   10425
   Begin VB.Frame Frame5 
      Height          =   975
      Left            =   3780
      TabIndex        =   21
      Top             =   120
      Width           =   6495
      Begin VB.Label Label7 
         Caption         =   "Veiculos que entraram e saírão"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   180
         TabIndex        =   22
         Top             =   300
         Width           =   6075
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Configurações"
      Height          =   1455
      Left            =   240
      TabIndex        =   11
      Top             =   120
      Width           =   2295
      Begin VB.TextBox Qvagas 
         Height          =   315
         Left            =   1500
         TabIndex        =   16
         Text            =   "40"
         Top             =   960
         Width           =   500
      End
      Begin VB.TextBox QEntradas 
         Height          =   315
         Left            =   1500
         TabIndex        =   13
         Top             =   240
         Width           =   500
      End
      Begin VB.TextBox QSaidas 
         Height          =   315
         Left            =   1500
         TabIndex        =   12
         Top             =   600
         Width           =   500
      End
      Begin VB.Label Label10 
         Caption         =   "Qtde de Vagas:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1020
         Width           =   1155
      End
      Begin VB.Label Label8 
         Caption         =   "Qtde de Entradas:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Qtde de Saída:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   660
         Width           =   1155
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   435
      Left            =   2700
      TabIndex        =   10
      Top             =   1140
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   435
      Left            =   2700
      TabIndex        =   9
      Top             =   240
      Width           =   735
   End
   Begin VB.Frame Frame3 
      Caption         =   "Saidas"
      Height          =   5460
      Left            =   6540
      TabIndex        =   4
      Top             =   2040
      Width           =   3615
      Begin MSComctlLib.ListView ListView3 
         Height          =   4000
         Left            =   120
         TabIndex        =   5
         Top             =   1260
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   7064
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Veículo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Saída"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "Intervalo:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Qtde veiculo:"
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   780
         Width           =   1635
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Estacionamento"
      Height          =   5460
      Left            =   4080
      TabIndex        =   2
      Top             =   2040
      Width           =   2355
      Begin MSComctlLib.ListView ListView2 
         Height          =   4005
         Left            =   120
         TabIndex        =   3
         Top             =   1260
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   7064
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Veiculos"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Intervalo:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Qtde veiculos:"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Entradas"
      Height          =   5460
      Left            =   180
      TabIndex        =   0
      Top             =   2040
      Width           =   3735
      Begin MSComctlLib.ListView ListView1 
         Height          =   4005
         Left            =   180
         TabIndex        =   1
         Top             =   1260
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   7064
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Veículo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Entrada"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Intervalo:"
         Height          =   255
         Left            =   180
         TabIndex        =   18
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Qtde Veiculos: "
         Height          =   375
         Left            =   180
         TabIndex        =   7
         Top             =   780
         Width           =   1395
      End
   End
   Begin VB.Timer Timer_saida 
      Left            =   8760
      Top             =   1680
   End
   Begin VB.Timer Timer_estacionado 
      Left            =   5460
      Top             =   1560
   End
   Begin VB.Timer Timer_Entrada 
      Left            =   2700
      Top             =   1680
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim F_in() As Collection
Dim F_stay As Collection
Dim F_out() As Collection
Dim I As Integer, y As Integer
Dim qtde_vagas As Integer
Dim qtde_veic As Integer
Dim qtde_veic_in As Integer
Dim qtde_veic_out As Integer
Dim qtde_veic_all As Integer
Dim num_veic As Integer
Dim indice_entrada As Integer
Dim indice_saida As Integer
Dim indice_entrada_estacionamento As Integer

Private Sub Command1_Click()
Dim I As Integer
Dim X As ListItem
Dim nroEntradas As Integer, contador As Integer
    
    If IsNumeric(QEntradas) And IsNumeric(QSaidas) Then
        
        qtde_vagas = CInt(Qvagas.Text)
        qtde_veic = 0
        qtde_veic_in = 0
        qtde_veic_out = 0
        qtde_veic_all = 0
        
        Dim entradas() As Collection
        nroEntradas = 5
        ReDim entradas(nroEntradas)
        For contador = 0 To nroEntradas - 1
            Set entradas(contador) = New Collection
        Next
        
        ReDim F_out(CInt(QSaidas.Text))
        For contador = 0 To QSaidas - 1
            Set F_out(contador) = New Collection
        Next
        
        Set F_stay = New Collection
        
        ReDim F_in(CInt(QEntradas.Text))
        For contador = 0 To QEntradas - 1
            Set F_in(contador) = New Collection
        Next
                
        ListView1.ListItems.Clear
        ListView2.ListItems.Clear
        ListView3.ListItems.Clear
        
        indice_entrada = 1
        indice_saida = 1
        indice_entrada_estacionamento = 0
        
        Timer_Entrada.Enabled = True
        Timer_estacionado.Enabled = True
        Timer_saida.Enabled = True
        
        Timer_Entrada.Interval = 1000
        Timer_estacionado.Interval = 10000
        Timer_saida.Interval = 5000
    End If
End Sub

Private Sub Command2_Click()
    Timer_Entrada.Interval = 0
    Timer_estacionado.Interval = 0
    Timer_saida.Interval = 0
    
    Timer_Entrada.Enabled = False
    Timer_estacionado.Enabled = False
    Timer_saida.Enabled = False
End Sub

Private Sub Form_Load()
Dim I As Integer
Dim X As ListItem
Dim nroEntradas As Integer, contador As Integer

    Timer_Entrada.Interval = 0
    Timer_estacionado.Interval = 0
    Timer_saida.Interval = 0
    
    Timer_Entrada.Enabled = False
    Timer_estacionado.Enabled = False
    Timer_saida.Enabled = False
End Sub

Private Sub Timer_Entrada_Timer()
Dim tempo1a5 As Integer
Dim X As ListItem

    'Transfere os carros da fila de entrada para a fila de estacionados
    tempo1a5 = (Rnd() * 5) + 1
    Timer_Entrada.Interval = 1000 * tempo1a5
    Label1.Caption = "Intervalo: " & tempo1a5 & " Segundos"
    
    If (ListView1.ListItems.Count > 0) And (F_stay.Count < qtde_vagas) Then
        F_stay.Add F_in(indice_entrada_estacionamento).Item(1)
        F_in(indice_entrada_estacionamento).Remove (1)
        indice_entrada_estacionamento = indice_entrada_estacionamento + 1
        
        If (indice_entrada_estacionamento >= QEntradas) Then
            indice_entrada_estacionamento = 0
        End If
        
        ListView2.ListItems.Clear
        
        For I = 1 To F_stay.Count
            Set X = Me.ListView2.ListItems.Add(, , F_stay(I))
        Next
        qtde_veic = qtde_veic + 1
        Label4.Caption = "Qtde Veiculos: " & qtde_veic
    
        ListView1.ListItems.Clear
        For y = 0 To QEntradas - 1
            If Not (IsNull(F_in(y).Count)) Then
                For I = 1 To F_in(y).Count
                    Set X = Me.ListView1.ListItems.Add(, , F_in(y).Item(I))
                    X.SubItems(1) = y
                Next
            End If
        Next
        qtde_veic_in = qtde_veic_in - 1
        Label5.Caption = "Qtde Veiculos: " & qtde_veic_in
        
        DoEvents
    End If
End Sub

Private Sub Timer_estacionado_Timer()
Dim tempo10a30 As Integer
Dim X As ListItem

    'Transfere os carros estacionados para a fila de saída
    tempo10a30 = (Rnd() * 20) + 10
    Timer_estacionado.Interval = 1000 * tempo10a30
    Label2.Caption = "Intervalo: " & tempo10a30 & " Segundos"

    If F_stay.Count > 0 Then
        F_out(indice_saida - 1).Add F_stay(1)
        F_stay.Remove (1)
        ListView3.ListItems.Clear
        For y = 0 To QSaidas - 1
            If Not (IsNull(F_out(y).Count)) Then
                For I = 1 To F_out(y).Count
                    If Not (IsNull(F_out(y).Item(I))) Then
                        Set X = Me.ListView3.ListItems.Add(, , F_out(y).Item(I))
                        X.SubItems(1) = y
                    End If
                Next
            End If
        Next
        qtde_veic = qtde_veic - 1
        Label4.Caption = "Qtde Veiculos: " & qtde_veic
        indice_saida = indice_saida + 1
        If indice_saida > QSaidas Then
            indice_saida = 1
        End If
        If indice_saida = 0 Then
            indice_saida = 1
        End If
    
        ListView2.ListItems.Clear
        For I = 1 To F_stay.Count
            Set X = Me.ListView2.ListItems.Add(, , F_stay.Item(I))
        Next
        qtde_veic_out = qtde_veic_out + 1
        Label6.Caption = "Qtde Veiculos: " & qtde_veic_out
        
        DoEvents
    End If
End Sub

Private Sub Timer_saida_Timer()
Dim tempo3a5 As Integer
Dim X As ListItem

    tempo3a5 = (Rnd() * 2) + 3
    Timer_saida.Interval = 1000 * tempo3a5
    Label3.Caption = "Intervalo: " & tempo3a5 & " Segundos"
    
    'Adiciona veicula na fila de entrada
    num_veic = num_veic + 1
    F_in(indice_entrada - 1).Add "Carro " & num_veic
    ListView1.ListItems.Clear
    For y = 0 To QEntradas - 1
        If Not (IsNull(F_in(y).Count)) Then
            For I = 1 To F_in(y).Count
                If Not (IsNull(F_in(y).Item(I))) Then
                    Set X = Me.ListView1.ListItems.Add(, , F_in(y).Item(I))
                    X.SubItems(1) = y
                End If
            Next
        End If
    Next
    
    indice_entrada = indice_entrada + 1
    If indice_entrada > QEntradas Then
        indice_entrada = 1
    End If
    
    qtde_veic_in = qtde_veic_in + 1
    Label5.Caption = "Qtde Veiculos: " & qtde_veic_in

    'Remove veiculo da fila de saída
    If F_out(indice_saida - 1).Count > 0 Then
        F_out(indice_saida - 1).Remove (1)
    
        ListView3.ListItems.Clear
        For y = 0 To QSaidas - 1
            If Not (IsNull(F_out(y).Count)) Then
                For I = 1 To F_out(y).Count
                    If Not (IsNull(F_out(y).Item(I))) Then
                        Set X = Me.ListView3.ListItems.Add(, , F_out(y).Item(I))
                        X.SubItems(1) = y
                    End If
                Next
            End If
        Next
        
        If indice_saida > 1 Then
            indice_saida = indice_saida - 1
        Else
            indice_saida = 1
        End If
        
        qtde_veic_out = qtde_veic_out - 1
        Label6.Caption = "Qtde Veiculos: " & qtde_veic_out
        
        qtde_veic_all = qtde_veic_all + 1
        Label7.Caption = "Veiculos que entraram e saírão: " & qtde_veic_all
    End If
    DoEvents
End Sub
