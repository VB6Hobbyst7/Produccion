VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmMntIndiceVac 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Indice VAC: Mantenimiento"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5700
   Icon            =   "frmMntIndiceVac.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   5700
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid dg 
      Height          =   3105
      Left            =   90
      TabIndex        =   7
      Top             =   120
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   5477
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   19
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "dIndiceVac"
         Caption         =   "FECHA"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   3
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "nIndiceVac"
         Caption         =   "VALOR"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00000000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   2654.929
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2220.094
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraControl 
      Height          =   555
      Left            =   90
      TabIndex        =   0
      Top             =   3210
      Width           =   5475
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   360
         Left            =   120
         TabIndex        =   6
         Top             =   150
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         CausesValidation=   0   'False
         Height          =   360
         Left            =   1200
         TabIndex        =   5
         Top             =   150
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         CausesValidation=   0   'False
         Height          =   360
         Left            =   4140
         TabIndex        =   4
         Top             =   150
         Width           =   1215
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   360
         Left            =   2280
         TabIndex        =   3
         Top             =   150
         Width           =   1065
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "&Modificar"
         Height          =   360
         Left            =   1200
         TabIndex        =   2
         Top             =   150
         Width           =   1065
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo"
         Height          =   360
         Left            =   120
         TabIndex        =   1
         Top             =   150
         Width           =   1065
      End
   End
   Begin VB.Frame fraDatos 
      Height          =   585
      Left            =   90
      TabIndex        =   8
      Top             =   2610
      Visible         =   0   'False
      Width           =   5475
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
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
         Left            =   3000
         TabIndex        =   10
         Top             =   180
         Width           =   2205
      End
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   315
         Left            =   330
         TabIndex        =   9
         Top             =   180
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
   End
End
Attribute VB_Name = "frmMntIndiceVac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSql As String
Dim rs   As ADODB.Recordset
Dim lNuevo As Boolean
Dim oCon As DConecta

Private Sub MuestraBotones(plActiva As Boolean)
fraDatos.Visible = plActiva
CmdAceptar.Visible = plActiva
CmdCancelar.Visible = plActiva
cmdNuevo.Visible = Not plActiva
cmdModificar.Visible = Not plActiva
cmdEliminar.Visible = Not plActiva
If plActiva Then
   dg.Height = dg.Height - 600
Else
   dg.Height = dg.Height + 600
End If
End Sub

Private Function ValidaDatos() As Boolean
ValidaDatos = False
If ValidaFecha(txtFecha) <> "" Then
   MsgBox "Fecha no válida", vbInformation, "¡Aviso!"
   txtFecha.SetFocus
   Exit Function
End If
If nVal(txtImporte) = 0 Then
   MsgBox "No se ingreso Importe de Indice", vbInformation, "¡Aviso!"
   txtImporte.SetFocus
   Exit Function
End If
ValidaDatos = True
End Function

Private Sub cmdAceptar_Click()
Dim nPos
If Not ValidaDatos() Then
   Exit Sub
End If
If MsgBox(" ¿ Seguro que desea Grabar datos ? ", vbQuestion + vbYesNo, "¡Confirmación!") = vbNo Then
   Exit Sub
End If
If lNuevo Then
    ' Verificamos si existe indice para ese dia SIGA 05052006
    Dim RVeri As ADODB.Recordset
    sSql = "Select * From IndiceVac WHERE dIndiceVac = '" & Format(txtFecha, gsFormatoFecha) & "'"
    oCon.Ejecutar sSql
    Set RVeri = oCon.CargaRecordSet(sSql, adLockOptimistic)
        
    If RVeri.RecordCount > 0 Then
        
        MsgBox "Ya existe Indice para el Dia", vbInformation, "Error al Grabar Indice"
        Exit Sub
        
    Else
            sSql = "INSERT IndiceVac (dIndiceVac, nIndiceVac) " _
                 & "VALUES ('" & Format(txtFecha, gsFormatoFecha) & "', " & CDec(Format(txtImporte, "#0.00######")) & ")"
            oCon.Ejecutar sSql
    End If
    
    RVeri.Close
    Set RVeri = Nothing
  
Else
   sSql = "UPDATE IndiceVac SET nIndiceVac = " & CDec(Format(txtImporte, "#0.00######")) & " WHERE dIndiceVac = '" & Format(txtFecha, gsFormatoFecha) & "'"
   oCon.Ejecutar sSql
End If
MuestraBotones False
CargaDatos
rs.Find "dIndiceVac = '" & txtFecha & "'", , adSearchForward, 0
dg.SetFocus
End Sub

Private Sub cmdCancelar_Click()
MuestraBotones False
dg.SetFocus
End Sub

Private Sub cmdEliminar_Click()
Dim nPos
If MsgBox(" ¿ Seguro que desea Eliminar datos ? ", vbQuestion + vbYesNo, "¡Confirmación") = vbYes Then
   nPos = rs.Bookmark
   sSql = "DELETE IndiceVac WHERE dIndiceVac = '" & Format(rs!dIndiceVac, gsFormatoFecha) & "'"
   oCon.Ejecutar sSql
   
   CargaDatos
   If nPos > rs.Bookmark Then
      rs.MoveLast
   Else
      rs.Bookmark = nPos
   End If
   rs.MoveLast
   dg.SetFocus
End If
End Sub

Private Sub cmdModificar_Click()
lNuevo = False
MuestraBotones True
txtFecha.Enabled = False
txtFecha = rs!dIndiceVac
txtImporte = Format(rs!nIndiceVac, "#.00######")
txtImporte.SetFocus
End Sub

Private Sub cmdNuevo_Click()
lNuevo = True
MuestraBotones True
txtFecha.Enabled = True
txtFecha.SetFocus
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
CentraForm Me
Set oCon = New DConecta
oCon.AbreConexion
CargaDatos
If Not rs.EOF Then
   rs.MoveFirst
End If
End Sub

Private Sub CargaDatos()
sSql = "SELECT * FROM INDICEVAC Order By dIndiceVac Desc"
Set rs = oCon.CargaRecordSet(sSql)
Set dg.DataSource = rs
End Sub

Private Sub Form_Unload(Cancel As Integer)
oCon.CierraConexion
Set oCon = Nothing
End Sub

Private Sub txtFecha_GotFocus()
fEnfoque txtFecha
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtImporte.SetFocus
End If
End Sub

Private Sub txtImporte_GotFocus()
fEnfoque txtImporte
End Sub

Private Sub txtImporte_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtImporte, KeyAscii, 16, 8)
If KeyAscii = 13 Then
   CmdAceptar.SetFocus
End If
End Sub
