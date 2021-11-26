VERSION 5.00
Begin VB.Form frmBuscaPersFioncodes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Busca Persona - Foncodes"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   Icon            =   "frmBuscaPersFoncodes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin SICMACT.FlexEdit fePersonas 
      Height          =   1755
      Left            =   1935
      TabIndex        =   9
      Top             =   795
      Width           =   5520
      _ExtentX        =   9737
      _ExtentY        =   3096
      Cols0           =   3
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "-Nombre-Doc Ident"
      EncabezadosAnchos=   "200-3800-1400"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X"
      ListaControles  =   "0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-C"
      FormatosEdit    =   "0-0-0"
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   195
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.Frame frabusca 
      Caption         =   "Buscar por ...."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   915
      Left            =   60
      TabIndex        =   4
      Top             =   195
      Width           =   1725
      Begin VB.OptionButton optOpcion 
         Caption         =   "Nº Docu&mento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   90
         TabIndex        =   6
         Top             =   585
         Width           =   1575
      End
      Begin VB.OptionButton optOpcion 
         Caption         =   "A&pellido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   90
         TabIndex        =   5
         Top             =   285
         Value           =   -1  'True
         Width           =   1200
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   270
      TabIndex        =   3
      Top             =   1575
      Width           =   1230
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   270
      TabIndex        =   2
      Top             =   1980
      Width           =   1230
   End
   Begin VB.TextBox txtDocPer 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1920
      MaxLength       =   15
      TabIndex        =   1
      Tag             =   "3"
      Top             =   390
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtNomPer 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1935
      TabIndex        =   0
      Tag             =   "1"
      Top             =   390
      Width           =   3990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ingrese Dato a Buscar :"
      Height          =   195
      Left            =   1920
      TabIndex        =   8
      Top             =   120
      Width           =   1680
   End
   Begin VB.Label LblDoc 
      Height          =   195
      Left            =   4005
      TabIndex        =   7
      Top             =   495
      Visible         =   0   'False
      Width           =   3495
   End
End
Attribute VB_Name = "frmBuscaPersFioncodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As Recordset

Private Sub CmdAceptar_Click()
Dim oDatos As DCapFideicomiso

Set oDatos = New DCapFideicomiso
    
    Set rs = oDatos.dGetCuentasFoncodes(fePersonas.TextMatrix(fePersonas.Row, 1))
    
    If Not rs.EOF And Not rs.BOF Then
        Do While Not rs.EOF
            frmCapFoncodes.LstCred.AddItem rs!cCtaCod
            
            rs.MoveNext
        Loop
    End If
    
    Set rs = Nothing
    Set oDatos = Nothing
    
    Unload Me
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub fePersonas_Click()
    txtNomPer = fePersonas.TextMatrix(fePersonas.Row, 1)
End Sub

Private Sub fePersonas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtNomPer = fePersonas.TextMatrix(fePersonas.Row, 1)
        CmdAceptar.SetFocus
    End If
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Private Sub optOpcion_Click(Index As Integer)
    
    If optOpcion(0).value = 1 Then
        txtNomPer.Visible = True
        txtDocPer.Visible = False
        txtNomPer.SetFocus
    Else
        txtNomPer.Visible = False
        txtDocPer.Visible = True
        txtDocPer.SetFocus
    End If

End Sub

Private Sub txtDocPer_KeyPress(KeyAscii As Integer)
Dim rs As Recordset
Dim oDatos As DCapFideicomiso

    If KeyAscii = 13 Then
        If txtDocPer <> "" Then
            Set oDatos = New DCapFideicomiso
            
            Set rs = oDatos.dGetDatosPersona(Trim(txtNomPer))
            
            Do While Not rs.EOF
                fePersonas.AdicionaFila
                fePersonas.TextMatrix(fePersonas.Rows - 1, 1) = rs!cNombre
                fePersonas.TextMatrix(fePersonas.Rows - 1, 2) = rs!cNumDocId
                rs.MoveNext
            Loop
            
            Set rs = Nothing
            Set oDatos = Nothing
        End If
    End If
    
End Sub

Private Sub txtNomPer_KeyPress(KeyAscii As Integer)
Dim rs As Recordset

Dim oDatos As DCapFideicomiso

    If KeyAscii = 13 Then
        If txtNomPer <> "" Then
            Set oDatos = New DCapFideicomiso
            
            Set rs = oDatos.dGetDatosPersonaFoncodes(Trim(txtNomPer))
            fePersonas.Clear
            fePersonas.FormaCabecera
        
            Do While Not rs.EOF
                fePersonas.AdicionaFila
                fePersonas.TextMatrix(fePersonas.Rows - 1, 1) = rs!cNombre
                fePersonas.TextMatrix(fePersonas.Rows - 1, 2) = rs!cNumDocId
                rs.MoveNext
            Loop
            
            Set rs = Nothing
            Set oDatos = Nothing
            
            fePersonas.SetFocus
            
        Else
            txtDocPer.SetFocus
        End If
    Else
        KeyAscii = Letras(KeyAscii)
    End If
    
End Sub

Public Sub Inicio(Optional ByVal psNomPers As String)

    Me.Show 1

End Sub


