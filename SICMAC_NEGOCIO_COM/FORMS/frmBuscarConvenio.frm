VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmBuscarConvenio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Buscar Convenio"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9750
   Icon            =   "frmBuscarConvenio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid dgConvenios 
      Height          =   1695
      Left            =   2730
      TabIndex        =   7
      Top             =   945
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   2990
      _Version        =   393216
      AllowUpdate     =   0   'False
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   1
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   " "
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "cPersnombre"
         Caption         =   "Empresa"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "cNombreConvenio"
         Caption         =   "Convenio"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "cPersCod"
         Caption         =   "Cod Empresa"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "cCodConvenio"
         Caption         =   "Cod Convenio"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "nTipo"
         Caption         =   "Tipo Convenio"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "cDescripcion"
         Caption         =   "Descripcion"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         Size            =   222
         BeginProperty Column00 
            ColumnWidth     =   2700.284
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4004.788
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtDatosBuscar 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      TabIndex        =   1
      Top             =   480
      Width           =   6780
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
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
      Left            =   525
      TabIndex        =   5
      Top             =   1920
      Width           =   1530
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
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
      Left            =   525
      TabIndex        =   4
      Top             =   1440
      Width           =   1530
   End
   Begin VB.Frame Frame1 
      Caption         =   "Buscar por"
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
      Height          =   1125
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   2430
      Begin VB.OptionButton cboDescripcion 
         Caption         =   "Nombre &Convenio"
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
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   2190
      End
      Begin VB.OptionButton cboNombre 
         Caption         =   "Nombre &Empresa"
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
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   2055
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Ingrese dato a buscar: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   120
      Width           =   2190
   End
End
Attribute VB_Name = "frmBuscarConvenio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim R As ADODB.Recordset
Dim bRespuesta As Boolean

Public Function Inicio() As Recordset
    
    Me.Show 1
    
    Set Inicio = R
    Set R = Nothing
    
End Function

Private Sub cboDescripcion_Click()
    txtDatosBuscar.Text = Empty
    txtDatosBuscar.SetFocus
End Sub

Private Sub cboNombre_Click()
    txtDatosBuscar.Text = Empty
    txtDatosBuscar.SetFocus
End Sub

Private Sub cmdAceptar_Click()

   If R Is Nothing Then
        MsgBox "Seleccione un Cliente", vbInformation, "Aviso"
        Exit Sub
   Else
        If R.RecordCount = 0 Then
            MsgBox "Seleccione un Cliente", vbInformation, "Aviso"
            Exit Sub
        End If
   End If
       
   R.Filter = "cCodConvenio = " & R!cCodConvenio
       
   Screen.MousePointer = 0
   bRespuesta = True
   Unload Me
   
End Sub

Private Sub cmdCancelar_Click()
    Set R = Nothing
    'Set Persona = Nothing
    Unload Me
End Sub

Private Sub Form_Load()
     dgConvenios.MarqueeStyle = dbgHighlightRow
     bRespuesta = False
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If Not bRespuesta Then
        Set R = Nothing
    End If
    
End Sub

Private Sub txtDatosBuscar_GotFocus()
    fEnfoque txtDatosBuscar
End Sub

Private Sub txtDatosBuscar_KeyPress(KeyAscii As Integer)

Dim ClsServicioRecaudo As COMDCaptaServicios.DCOMServicioRecaudo

If KeyAscii = 13 Then
    If Len(Trim(txtDatosBuscar.Text)) = 0 Then
         MsgBox "Falta Ingresar los datos para la busqueda", vbInformation, "Aviso"
         Exit Sub
    End If
    Screen.MousePointer = 11
    Set ClsServicioRecaudo = New COMDCaptaServicios.DCOMServicioRecaudo
    If cboNombre.value Then
        Set R = ClsServicioRecaudo.getBuscarConvenio(txtDatosBuscar.Text)
    Else
        Set R = ClsServicioRecaudo.getBuscarConvenio(, , txtDatosBuscar.Text)
    End If
    Set dgConvenios.DataSource = R
    dgConvenios.Refresh
    Screen.MousePointer = 0
    
    If R.RecordCount = 0 Then
         MsgBox "No se Encontraron Datos", vbInformation, "Aviso"
         txtDatosBuscar.SetFocus
         cmdAceptar.Default = False
    Else
         cmdAceptar.Default = True
         dgConvenios.SetFocus
    End If
Else
     KeyAscii = Letras(KeyAscii)
     cmdAceptar.Default = False
End If

End Sub

