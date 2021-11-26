VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCajaArqueoVentBovExt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Arqueos: Extornos"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10950
   Icon            =   "frmCajaArqueoVentBovExt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   10950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   9600
      TabIndex        =   5
      Top             =   3480
      Width           =   1170
   End
   Begin VB.CommandButton cmdExtornar 
      Caption         =   "Extornar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8400
      TabIndex        =   4
      Top             =   3480
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.TextBox txtGlosa 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   840
      TabIndex        =   2
      Top             =   3480
      Width           =   7335
   End
   Begin VB.Frame Frame1 
      Caption         =   " Listado de Arqueos "
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10695
      Begin MSDataGridLib.DataGrid DBGrdArqueos 
         Height          =   2655
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   10395
         _ExtentX        =   18336
         _ExtentY        =   4683
         _Version        =   393216
         AllowUpdate     =   0   'False
         ColumnHeaders   =   -1  'True
         HeadLines       =   2
         RowHeight       =   17
         RowDividerStyle =   4
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
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "nMovNro"
            Caption         =   "nMovNro"
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
            DataField       =   "cTipo"
            Caption         =   "Tipo"
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
            DataField       =   "cLugar"
            Caption         =   "Lugar"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "cPersArqueado"
            Caption         =   "Personal Arqueado"
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
            DataField       =   "cArqueadorAudit"
            Caption         =   "Arqueador/Auditor"
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
            DataField       =   "cResult"
            Caption         =   "Resultado"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "cGlosa"
            Caption         =   "Glosa"
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
         BeginProperty Column07 
            DataField       =   "nTipoArqueo"
            Caption         =   "nTipoArqueo"
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
            SizeMode        =   1
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            Size            =   800
            BeginProperty Column00 
               Object.Visible         =   0   'False
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1275.024
            EndProperty
            BeginProperty Column02 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1275.024
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2505.26
            EndProperty
            BeginProperty Column04 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   2505.26
            EndProperty
            BeginProperty Column05 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column06 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   2505.26
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   0
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Glosa : "
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   3525
      Width           =   735
   End
End
Attribute VB_Name = "frmCajaArqueoVentBovExt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'** Nombre : frmOpeArqueoVentBovExt
'** Descripción : Extorno de Procesos de Arqueos de Ventanillas y Bóvedas del dia creado segun RFC081-2012
'** Creación : JUEZ, 20120813 09:00:00 AM
'********************************************************************

Option Explicit
Dim oVisto As frmVistoElectronico
Dim bResultadoVisto As Boolean
Dim oCajaGen As COMNCajaGeneral.NCOMCajaGeneral
Dim RS As ADODB.Recordset

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdExtornar_Click()
    Dim oContFunct As COMNContabilidad.NCOMContFunciones
    Set oCajaGen = New COMNCajaGeneral.NCOMCajaGeneral
    Set oContFunct = New COMNContabilidad.NCOMContFunciones
    
    If Trim(txtGlosa.Text) = "" Then
        MsgBox "Necesita escribir la glosa para proceder con el extorno", vbInformation, "Aviso"
        txtGlosa.SetFocus
        Exit Sub
    End If
    
    Dim lnMovNro As Long
    Dim lsMovNro As String
    Dim nNroRegistros As Integer 'RIRO20140710 ERS072
     
    lnMovNro = DBGrdArqueos.Columns(0)
    
    ' *** RIRO20140710 ERS072
        If Val(DBGrdArqueos.Columns(7)) = 3 Then
            If oCajaGen.ObtenerNroExtornos(lnMovNro, nNroRegistros) Then
                If nNroRegistros > 1 Then
                    MsgBox "El número máximo de extornos fue utilizado", vbExclamation, "Aviso"
                    Exit Sub
                End If
            Else
                MsgBox "Se presento un error durante el proceso de extorno", vbCritical, "Aviso"
                Exit Sub
            End If
        End If
    ' *** END RIRO
    
    If lnMovNro <> 0 Then
        
        Set oVisto = New frmVistoElectronico
        bResultadoVisto = oVisto.Inicio(6)
        If Not bResultadoVisto Then
            Exit Sub
        End If
        
        If MsgBox("Está seguro de extornar el proceso de arqueo seleccionado? ", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
         
        lsMovNro = oContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
         
        Call oCajaGen.ExtornaProcesoArqueo(lnMovNro, lsMovNro, txtGlosa.Text, DBGrdArqueos.Columns(1))
        MsgBox "Se ha extornado el arqueo seleccionado", vbInformation, "Aviso"
        Me.txtGlosa.Text = ""
        Call Form_Load
    End If
End Sub

Private Sub Form_Load()
    Set oCajaGen = New COMNCajaGeneral.NCOMCajaGeneral
    
    Set RS = oCajaGen.ObtieneListaArqueosRealizados(gsCodAge, gdFecSis)
    Set DBGrdArqueos.DataSource = RS
    DBGrdArqueos.Refresh
    
    If RS.RecordCount = 0 Then
      MsgBox "No se Encontraron Datos Registrados", vbInformation, "Aviso"
      cmdExtornar.Visible = False
      txtGlosa.Enabled = False
    Else
      cmdExtornar.Visible = True
      txtGlosa.Enabled = True
    End If
End Sub

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdExtornar.SetFocus
    End If
End Sub
