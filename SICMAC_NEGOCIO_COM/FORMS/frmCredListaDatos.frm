VERSION 5.00
Begin VB.Form frmCredListaDatos 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3645
   Icon            =   "frmCredListaDatos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   3645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstLista 
      Height          =   1425
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.TextBox txtTexto 
      Height          =   1635
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   3255
   End
   Begin SICMACT.FlexEdit feDatos1 
      Height          =   1635
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   1931
      Cols0           =   3
      HighLight       =   1
      EncabezadosNombres=   "-Cargo-Tipo"
      EncabezadosAnchos=   "300-2400-880"
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X"
      ListaControles  =   "0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-C"
      FormatosEdit    =   "0-1-0"
      lbUltimaInstancia=   -1  'True
      TipoBusqueda    =   3
      ColWidth0       =   300
      RowHeight0      =   300
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmCredListaDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fbCargado As Boolean
Dim fbEdita As Boolean

Public Sub Inicio(ByVal psTitulo As String, ByVal rsDatos As ADODB.Recordset, _
                  Optional ByVal rsLista As ADODB.Recordset = Nothing, Optional ByVal pnTipoLista As Integer = 0, _
                  Optional ByVal pnTitulosFlex As Variant = Nothing)
    Dim i As Integer, j As Integer
    Me.Caption = psTitulo
    fbCargado = False
    If pnTipoLista = 1 Then 'ListBox
        Me.Width = 3600
        lstLista.Visible = True
        lstLista.Clear
        If Not rsDatos.EOF Then
            For i = 0 To rsLista.RecordCount - 1
                rsDatos.MoveFirst
                For j = 0 To rsDatos.RecordCount - 1
                    If Trim(rsLista.Fields(1)) = Trim(rsDatos.Fields(1)) Then
                        lstLista.AddItem rsLista.Fields(1) & " " & Trim(rsLista.Fields(0))
                        Exit For
                    End If
                    rsDatos.MoveNext
                Next j
                rsLista.MoveNext
            Next i
        Else
            MsgBox "No se puede obtener los datos", vbExclamation, "Aviso"
            Exit Sub
        End If
    ElseIf pnTipoLista = 2 Then 'FlexEdit 2 Filas
        Me.Width = 4305
        feDatos1.Visible = True
        Call LimpiaFlex(feDatos1)
        feDatos1.TextMatrix(0, 1) = pnTitulosFlex(0, 0)
        feDatos1.TextMatrix(0, 2) = pnTitulosFlex(0, 1)
        Do While Not rsDatos.EOF
            feDatos1.AdicionaFila
            i = feDatos1.row
            feDatos1.TextMatrix(i, 1) = rsDatos.Fields(2)
            feDatos1.TextMatrix(i, 2) = rsDatos.Fields(4)
            rsDatos.MoveNext
        Loop
        rsDatos.Close
        Set rsDatos = Nothing
    End If
    fbCargado = True
    Me.Show 1
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

'Private Sub lstLista_ItemCheck(iTem As Integer)
'    If fbCargado Then
'        If lstLista.Selected(iTem) Then
'            lstLista.Selected(iTem) = False
'        Else
'            lstLista.Selected(iTem) = True
'        End If
'    End If
'End Sub

Public Sub InicioTextBox(ByVal psTitulo As String, ByVal psTexto As String)
    Me.Caption = psTitulo
    Me.Width = 3600
    txtTexto.Visible = True
    txtTexto.Text = psTexto
    Me.Show 1
End Sub
