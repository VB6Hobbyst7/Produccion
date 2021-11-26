VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLogProSelReq 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Requerimientos de Usuario"
   ClientHeight    =   6450
   ClientLeft      =   1380
   ClientTop       =   1770
   ClientWidth     =   9705
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogProSelReq.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   9705
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7140
      TabIndex        =   13
      Top             =   6000
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   12
      Top             =   6000
      Width           =   1200
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   120
      TabIndex        =   0
      Top             =   -60
      Width           =   9465
      Begin VB.TextBox txtAgencia 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   900
         Width           =   2745
      End
      Begin VB.TextBox txtArea 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   900
         Width           =   5565
      End
      Begin VB.TextBox txtCargo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   600
         Width           =   8325
      End
      Begin VB.TextBox txtPersona 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   300
         Width           =   6585
      End
      Begin VB.CommandButton cmdPersona 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2310
         TabIndex        =   1
         Top             =   330
         Width           =   315
      End
      Begin VB.TextBox txtPersCod 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   300
         Width           =   1755
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Area"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   960
         Width           =   330
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Cargo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   660
         Width           =   420
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Usuario"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   360
         Width           =   540
      End
   End
   Begin TabDlg.SSTab sstReg 
      Height          =   4620
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   8149
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   670
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "     Requerimiento de Bienes / Servicios        "
      TabPicture(0)   =   "frmLogProSelReq.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "         Lista de Requerimientos                     "
      TabPicture(1)   =   "frmLogProSelReq.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MSFlexLista"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   120
         TabIndex        =   14
         Top             =   420
         Width           =   9195
         Begin VB.TextBox txtAnio 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2100
            MaxLength       =   4
            TabIndex        =   21
            Top             =   240
            Width           =   760
         End
         Begin VB.ComboBox cboMes 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2880
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   240
            Width           =   6165
         End
         Begin VB.CommandButton cmdAgregar 
            Caption         =   "Agregar"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            TabIndex        =   17
            Top             =   2100
            Width           =   1155
         End
         Begin VB.CommandButton cmdQuitar 
            Caption         =   "Quitar"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1320
            TabIndex        =   16
            Top             =   2100
            Width           =   1155
         End
         Begin VB.TextBox txtEdit 
            BackColor       =   &H00DDFFFE&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4560
            MaxLength       =   6
            TabIndex        =   15
            Top             =   1620
            Visible         =   0   'False
            Width           =   990
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlex 
            Height          =   1455
            Left            =   120
            TabIndex        =   18
            Top             =   600
            Width           =   8955
            _ExtentX        =   15796
            _ExtentY        =   2566
            _Version        =   393216
            BackColor       =   16777215
            Cols            =   7
            FixedCols       =   0
            ForeColorFixed  =   -2147483646
            BackColorBkg    =   16777215
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483633
            GridColorUnpopulated=   -2147483633
            FocusRect       =   0
            HighLight       =   2
            AllowUserResizing=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   7
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Mes de requerimiento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   20
            Top             =   300
            Width           =   1830
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlexLista 
         Height          =   3000
         Left            =   -74880
         TabIndex        =   11
         Top             =   540
         Width           =   9195
         _ExtentX        =   16219
         _ExtentY        =   5292
         _Version        =   393216
         BackColor       =   16777215
         Cols            =   6
         FixedCols       =   0
         ForeColorFixed  =   -2147483646
         BackColorBkg    =   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         GridColorUnpopulated=   -2147483633
         FocusRect       =   0
         HighLight       =   2
         ScrollBars      =   2
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   120
         TabIndex        =   22
         Top             =   2805
         Width           =   9195
         Begin VB.CommandButton cmdAgregarPers 
            Caption         =   "Agregar Persona"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   7920
            TabIndex        =   25
            Top             =   420
            Width           =   1155
         End
         Begin VB.CommandButton cmdQuitarPers 
            Caption         =   "Quitar Persona"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   7920
            TabIndex        =   24
            Top             =   1020
            Width           =   1155
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSPers 
            Height          =   1155
            Left            =   120
            TabIndex        =   23
            Top             =   405
            Width           =   7695
            _ExtentX        =   13573
            _ExtentY        =   2037
            _Version        =   393216
            BackColor       =   16777215
            Cols            =   6
            FixedCols       =   0
            ForeColorFixed  =   -2147483646
            BackColorBkg    =   16777215
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483633
            GridColorUnpopulated=   -2147483633
            FocusRect       =   0
            HighLight       =   2
            ScrollBars      =   2
            AllowUserResizing=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   6
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Personal del Requerimiento "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   195
            Width           =   2400
         End
      End
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "MenuReq"
      Visible         =   0   'False
      Begin VB.Menu mnuInfo 
         Caption         =   "Info del trámite "
      End
   End
   Begin VB.Menu mnuPersonas 
      Caption         =   "mnuPersonas"
      Visible         =   0   'False
      Begin VB.Menu mnuInsertar 
         Caption         =   "Insertar Persona"
      End
      Begin VB.Menu mnuQuitar 
         Caption         =   "Quitar Persona"
      End
   End
   Begin VB.Menu mnuEstado 
      Caption         =   "MenuEstado"
      Visible         =   0   'False
      Begin VB.Menu mnuEstadoAprobacion 
         Caption         =   "Estado de Aprobacion"
      End
   End
End
Attribute VB_Name = "frmLogProSelReq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mMes(1 To 12) As String, sSQL As String
Dim cRHAgeCod As String, cRHAreaCod As String, cRHCargoCod As String
Dim nEditable As Boolean, nReqActual As Long

'Private Function VerificaCantidades() As Boolean
'On Error GoTo VerificaCantidadesErr
'    Dim i As Integer
'    With MSFlex
'        i = 2
'        Do While i < .Rows
'            If Val(.TextMatrix(i, 5)) = 0 And Val(.TextMatrix(i, 0)) = 0 Then
'                VerificaCantidades = False
'                Exit Function
'            End If
'            i = i + 1
'        Loop
'        VerificaCantidades = True
'    End With
'    Exit Function
'VerificaCantidadesErr:
'    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
'End Function

Private Sub cboMes_Click()
    If (cboMes.ListIndex + 1) < Month(gdFecSis) Then cboMes.ListIndex = Month(gdFecSis) - 1
End Sub

Private Sub cmdGrabar_Click()
Dim oConn As DConecta, sSQL As String
Dim cProSelBSCod As String, i As Integer, n As Integer
Dim nProSelReq As Long, Rs As New ADODB.Recordset
Dim nItem As Integer
Dim cLogNro As String
Dim nAnio As Integer, cPersCod As String
Dim k As Integer, m As Integer

n = MSFlex.Rows - 1
For i = 1 To n
    If Not VerificaCantidad(i, True) Then
       MSFlex.row = i
       MSFlex.Col = 3
       MSFlex.SetFocus
       Exit Sub
    End If
Next

For i = 1 To n
    If Len(Trim(MSFlex.TextMatrix(i, 2))) > 0 And Len(Trim(MSFlex.TextMatrix(i, 6))) = 0 Then
       MsgBox "Debe ingresar un sustento para todos los requerimientos...", vbInformation, "Aviso"
       MSFlex.row = i
       MSFlex.Col = 6
       MSFlex.SetFocus
       Exit Sub
    End If
Next

nItem = 0
n = MSFlex.Rows - 1
For i = 1 To n
    If Len(MSFlex.TextMatrix(i, 2)) > 0 And Len(MSFlex.TextMatrix(i, 3)) > 0 Then
       nItem = nItem + 1
    End If
Next

If nItem = 0 Then
   MsgBox "Debe seleccionar al menos un Bien / Servicio..." + Space(10), vbInformation
   Exit Sub
End If

'Exit Sub
nItem = 0
nAnio = Year(gdFecSis)
Set oConn = New DConecta
If oConn.AbreConexion Then
   If MsgBox("¿ Está seguro de grabar ?" + Space(10), vbQuestion + vbYesNo, "Confirme") = vbYes Then
   
      cLogNro = GetLogMovNro
      cPersCod = IIf(Len(txtPersCod.Text) = 13, txtPersCod.Text, gsCodPersUser)
      '1º Deshabilitamos los requerimientos anteriores del usuario
'      sSQL = "UPDATE LogProSelReq SET nEstado = 0 WHERE cPersCod = '" & txtPersCod.Text & "'"
'      oConn.Ejecutar sSQL
'
'      sSQL = "UPDATE LogProSelReqDetalle SET nEstado = 0 WHERE cPersCod = '" & txtPersCod.Text & "'"
'      oConn.Ejecutar sSQL
'
      '2º Insertamos cabecera del requerimiento actual del usuario
      sSQL = "INSERT INTO LogProSelReq( nAnio, cPersCod, cRHCargoCod, cRHAreaCod, cRHAgeCod, cMovNro,  nMesEje) " & _
             "    VALUES (" & nAnio & ",'" & cPersCod & "','" & cRHCargoCod & "','" & cRHAreaCod & "','" & cRHAgeCod & "','" & cLogNro & "'," & cboMes.ListIndex + 1 & ") "
      oConn.Ejecutar sSQL
      
      '3º Hallamos ultima secuencia de los requerimientos
      nProSelReq = UltimaSecuenciaIdentidad("LogProSelReq")
      
      '---------------------------------------------------------------------------------
      
      m = MSPers.Rows - 1
      nItem = 0
      For i = 1 To n
          nItem = nItem + 1
          sSQL = "INSERT INTO LogProSelReqDetalle (nProSelReqNro,cPersCod,nAnio, nItem, cBSCod,nCantidad,cSustento) " & _
                 " VALUES (" & nProSelReq & ",'" & cPersCod & "'," & nAnio & "," & nItem & ",'" & MSFlex.TextMatrix(i, 2) & "'," & MSFlex.TextMatrix(i, 5) & ",'" & MSFlex.TextMatrix(i, 6) & "' )"
          oConn.Ejecutar sSQL
             
          For k = 1 To m
              If MSFlex.TextMatrix(i, 2) = MSPers.TextMatrix(k, 1) Then
                 sSQL = "INSERT INTO LogProSelReqPersonas (nProSelReqNro,cPersCod,nItem, cBSCod,cPersCodDestino) " & _
                        " VALUES (" & nProSelReq & ",'" & cPersCod & "'," & nItem & ",'" & MSFlex.TextMatrix(i, 2) & "','" & MSPers.TextMatrix(k, 2) & "')"
                 oConn.Ejecutar sSQL
              End If
          Next
      Next
      
      sSQL = " insert into LogProSelAprobacion (nProSelReqNro,cRHCargoCodAprobacion,nNivelAprobacion) " & _
             " select distinct " & nProSelReq & ",cRHCargoCodAprobacion,nNivelAprobacion " & _
             " from LogNivelAprobacion where cRHCargoCod = '" & cRHCargoCod & "' order by nNivelAprobacion "
      oConn.Ejecutar sSQL
      
      MsgBox "El requerimiento se ha grabado con éxito!" + Space(10), vbInformation
      txtPersCod = ""
      FormaFlex
      FlexPersonas
      cmdGrabar.Enabled = True
   End If
   oConn.CierraConexion
End If
End Sub

Private Sub cmdQuitar_Click()
Dim i As Integer
Dim k As Integer, nPapa As Integer

i = MSFlex.row
nPapa = i

If Len(Trim(MSFlex.TextMatrix(i, 2))) = 0 Or MSFlex.TextMatrix(i, 0) <> "" Then
   Exit Sub
End If

If MsgBox("¿ está seguro de quitar el elemento ?" + Space(10), vbQuestion + vbYesNo, "Confirme") = vbYes Then
   Do While i < MSFlex.Rows
        If MSFlex.Rows - 1 > 1 Then
           MSFlex.RemoveItem i
        Else
           'MSFlex.Clear          Quita las cabeceras
           For k = 0 To MSFlex.Cols - 1
                 MSFlex.TextMatrix(i, k) = ""
           Next
           MSFlex.RowHeight(i) = 8
        End If
        i = MSFlex.row
        If MSFlex.TextMatrix(i, 0) = "" Or Val(MSFlex.TextMatrix(i, 7)) <> nPapa Then Exit Sub
    Loop
End If
End Sub

Private Sub Form_Load()
Dim oAcceso As UAcceso
Set oAcceso = New UAcceso
CentraForm Me
FlexPersonas
txtanio = Year(Date)
GeneraMeses
cRHAgeCod = ""
cRHAreaCod = ""
cRHCargoCod = ""
txtAgencia.Text = ""
txtCargo.Text = ""
txtArea.Text = ""
txtPersCod.Text = gsCodPersUser
FormaFlex
If cboMes.ListCount > 0 Then cboMes.ListIndex = Month(gdFecSis) - 1
sstReg.Tab = 0
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdPersona_Click()
Dim X As UPersona
Set X = frmBuscaPersona.Inicio(True)

If X Is Nothing Then
    Exit Sub
End If

If Len(Trim(X.sPersNombre)) > 0 Then
   txtPersona.Text = X.sPersNombre
   txtPersCod = X.sPersCod
End If
End Sub

Sub DatosPersonal(vPersCod As String)
Dim oPlan As New DLogPlanAnual
Dim Rs As New ADODB.Recordset
Dim rn As New ADODB.Recordset

Dim i As Integer, k As Integer, nAprobado As Integer
Dim nNivel As Integer, nSector As Integer, nAnio As Integer
Dim cAnioMesActual As String

nReqActual = 0
cRHAgeCod = ""
cRHAreaCod = ""
cRHCargoCod = ""
nEditable = False

cmdAgregar.Enabled = False
cmdQuitar.Enabled = False
cmdGrabar.Visible = False

FormaFlex

'Set oConn = New DConecta
'sSQL = "select x.*,cCargo=coalesce(c.cRHCargoDescripcion,''),cArea=coalesce(a.cAreaDescripcion,''),cAgencia=coalesce(b.cAgeDescripcion,'') " & _
'       " from (select top 1 cRHCargoCodOficial as cRHCargoCod, cRHAreaCodOficial as cAreaCod, cRHAgenciaCodOficial as cAgeCod " & _
'       "  from RHCargos where cPersCod='" & vPersCod & "' order by dRHCargoFecha desc) x " & _
'       "  left outer join Areas a on x.cAreaCod = a.cAreaCod " & _
'       "  left outer join Agencias b on x.cAgeCod = b.cAgeCod " & _
'       "  left outer join RHCargosTabla c on x.cRHCargoCod = c.cRHCargoCod "

cAnioMesActual = CStr(Year(gdFecSis)) + Format(Month(gdFecSis), "00")
'If oConn.AbreConexion Then
   Set Rs = oPlan.AreaCargoAgencia(vPersCod, cAnioMesActual)
   'oConn.CierraConexion
   If Not Rs.EOF Then
      cRHAgeCod = Rs!cRHAgeCod
      cRHAreaCod = Rs!cRHAreaCod
      cRHCargoCod = Rs!cRHCargoCod
      txtAgencia.Text = Rs!cRHAgencia
      txtCargo.Text = Rs!cRHCargo
      txtArea.Text = Rs!cRHArea
      txtPersona.Text = Rs!cPersona

      'Verifica NIVELES DE APROBACION -----------------------------------------------
      Set rn = GetNivelesAprobacion(cRHCargoCod)
      If rn.State = 0 Then
         MsgBox "No se puede determinar niveles de aprobacion para: " + Space(10) + vbCrLf + txtCargo.Text + Space(10) + vbCrLf + txtArea.Text + Space(10), vbInformation
         cmdQuitar.Enabled = False
         cmdAgregar.Enabled = False
         cmdGrabar.Visible = False
         Exit Sub
      Else
         If rn.EOF And rn.BOF Then
            MsgBox "No se puede determinar niveles de aprobacion para: " + Space(10) + vbCrLf + txtCargo.Text + Space(10) + vbCrLf + txtArea.Text + Space(10), vbInformation
            cmdQuitar.Enabled = False
            cmdAgregar.Enabled = False
            cmdGrabar.Visible = False
            Exit Sub
         End If
      End If
      
     '--------------------------------------------------------------
      'Verifica si hay requerimientos grabados ----------------------
      '--------------------------------------------------------------
      nReqActual = 0
      nAprobado = 0
      Set Rs = oPlan.EstadoAprobacionRequerimiento(nAnio, vPersCod, 1, 1)
      If Not Rs.EOF Then
         nReqActual = Rs!nPlanReqNro
         nAprobado = Rs!nEstadoAprobacion
         'Si ya está aprobado |||||||||||||||||||||||||||||||||||||||
         If nAprobado = 1 Then
            nEditable = False
            cmdQuitar.Enabled = False
            cmdAgregar.Enabled = False
            cmdGrabar.Visible = False
            'lblReqNro.Caption = "Requerimiento Nº " + CStr(nReqActual)
            MsgBox "El requerimiento ya fue aprobado !" + Space(10), vbInformation
         Else
            nEditable = True
            cmdAgregar.Enabled = True
            cmdQuitar.Enabled = True
            cmdGrabar.Visible = True
         End If
      Else
         nEditable = True
         cmdAgregar.Enabled = True
         cmdQuitar.Enabled = True
         cmdGrabar.Visible = True
      End If
      Set Rs = Nothing
      ConsultarRequerimientos vPersCod
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set frmLogProSelReq = Nothing
End Sub

Private Sub mnuEstadoAprobacion_Click()
    With MSFlexLista
        If .TextMatrix(.row, 1) <> "+" And .TextMatrix(.row, 1) <> "-" Then Exit Sub
        frmLogPlanAnualInfo.PlanAnual 0, True, .TextMatrix(.row, 2)
    End With
End Sub

Private Sub mnuInsertar_Click()
    Dim X As UPersona, i As Integer
    Set X = frmBuscaPersona.Inicio(True)
    
  '  If X Is Nothing Then
  '      Exit Sub
  '  End If
    
  '  If Len(Trim(X.sPersNombre)) > 0 Then
  '      If VerificaPersona(X.sPersCod, MSFlex.TextMatrix(MSFlex.row, 2)) Then
  '          MsgBox "Persona ya fue Ingresada", vbInformation, "Aviso"
  '          Exit Sub
  '      End If
  '      If Not VerificaCantidad(Val(MSFlex.TextMatrix(MSFlex.row, 5)), MSFlex.TextMatrix(MSFlex.row, 2)) Then
  '          MsgBox "Nro de Personas debe Conincidir con la Cantidad", vbInformation, "Aviso"
  '          Exit Sub
  '      End If
  '      i = MSFlex.row
  '      With MSFlex
  '          If Not VerificarAreaAgencia(X.sPersCod) Then
  '              MsgBox "Debe ser de la misma area y agencia", vbInformation, "Aviso"
  '              Exit Sub
  '          End If
' '           .Rows = .Rows + 1
  '          .Col = 1
  '          .CellFontBold = True
  '          .CellFontSize = 10
  '          .CellAlignment = 4
  '          .TextMatrix(i, 1) = "-"
  '          InsRow MSFlex, .Rows
  '          .TextMatrix(.Rows - 1, 0) = X.sPersCod
  '          .TextMatrix(.Rows - 1, 2) = .TextMatrix(i, 2)
  '          .TextMatrix(.Rows - 1, 3) = X.sPersNombre
  '          .TextMatrix(.Rows - 1, 7) = i
  '      End With
  '  End If
End Sub

Private Sub mnuQuitar_Click()
Dim i As Integer
Dim k As Integer

If MSFlex.TextMatrix(MSFlex.Rows - 1, 0) = "" Then Exit Sub

i = MSFlex.row
If MSFlex.TextMatrix(i, 0) = "" Then
   Exit Sub
End If

    If MsgBox("¿ está seguro de quitar el elemento ?" + Space(10), vbQuestion + vbYesNo, "Confirme") = vbNo Then Exit Sub
    If MSFlex.Rows - 1 > 1 Then
       MSFlex.RemoveItem i
    Else
       For k = 0 To MSFlex.Cols - 1
             MSFlex.TextMatrix(i, k) = ""
       Next
       MSFlex.RowHeight(i) = 8
    End If
End Sub

Private Sub MSFlex_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete And MSFlex.Col >= 4 And MSFlex.Col <= 15 And nEditable Then
   MSFlex.TextMatrix(MSFlex.row, MSFlex.Col) = ""
End If
End Sub

Private Sub MSFlex_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    PopupMenu mnuPersonas
End If
End Sub

Private Sub txtAnio_KeyPress(KeyAscii As Integer)
nKeyAscii = KeyAscii
KeyAscii = DigNumEnt(KeyAscii)
If nKeyAscii = 13 Then
   txtPersCod.SetFocus
End If
End Sub

Private Sub MSFlexLista_DblClick()
On Error GoTo MSFlexErr
    Dim i As Integer, bTipo As Boolean
    With MSFlexLista
        If Trim(.TextMatrix(.row, 1)) = "-" Then
           .TextMatrix(.row, 1) = "+"
           i = .row + 1
           bTipo = True
        ElseIf Trim(.TextMatrix(.row, 1)) = "+" Then
           .TextMatrix(.row, 1) = "-"
           i = .row + 1
           bTipo = False
        End If
        Do While i < .Rows
            If Trim(.TextMatrix(i, 1)) = "+" Or Trim(.TextMatrix(i, 1)) = "-" Then
                Exit Sub
            End If
            
            If bTipo Then
                .RowHeight(i) = 0
            Else
                .RowHeight(i) = 260
            End If
            i = i + 1
        Loop
    End With
Exit Sub
MSFlexErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Private Sub MSFlexLista_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuEstado
    End If
End Sub

Private Sub txtPersCod_Change()
If Len(txtPersCod) > 0 Then
    DatosPersonal IIf(Len(txtPersCod.Text) = 13, txtPersCod.Text, gsCodPersUser)
Else
   FormaFlex
   cmdGrabar.Enabled = False
   cmdAgregar.Enabled = False
   cmdQuitar.Enabled = False
   txtArea.Text = ""
   txtAgencia.Text = ""
   txtCargo.Text = ""
   txtPersona.Text = ""
End If
End Sub

Sub Totaliza()
Dim i As Integer, j As Integer, n As Integer
Dim nSuma As Currency
n = MSFlex.Rows - 1
For i = 1 To n
    nSuma = 0
    For j = 1 To 13
        nSuma = nSuma + VNumero(MSFlex.TextMatrix(i, j + 4))
    Next
    MSFlex.TextMatrix(i, 16) = nSuma
Next
End Sub

Private Sub cmdAgregar_Click()
Dim i As Integer, Rs As New ADODB.Recordset
        
i = MSFlex.Rows - 1
If i = 1 And Len(Trim(MSFlex.TextMatrix(i, 2))) = 0 Then
   i = 0
End If

frmLogProSelBSSelector.TodosConCheck True, True, False
Set Rs = frmLogProSelBSSelector.gvrs
If Rs.State <> 0 Then
   If Not Rs.EOF Then
      Do While Not Rs.EOF
         If Not YaEsta(Rs!cProSelBSCod) Then
            i = i + 1
            InsRow MSFlex, i
            MSFlex.TextMatrix(i, 2) = Rs!cProSelBSCod
            MSFlex.TextMatrix(i, 3) = Rs!cBSDescripcion
            MSFlex.TextMatrix(i, 4) = GetBSUnidadLog(Rs!cProSelBSCod)
            MSFlex.TextMatrix(i, 5) = "1"
            MSFlex.row = MSFlex.Rows - 1
            MSFlex.Col = 5
            frmLogProSelEspecificaciones.Inicio MSFlex.Left + MSFlex.CellLeft, MSFlex.Top + MSFlex.CellTop + 3600, "Sustento: "
            MSFlex.TextMatrix(i, 6) = frmLogProSelEspecificaciones.vpTexto
         End If
         Rs.MoveNext
      Loop
   End If
End If
End Sub

Private Sub MSFlex_DblClick()
Dim nFila As Integer
nFila = MSFlex.row
frmLogProSelEspecificaciones.Inicio MSFlex.Left + MSFlex.CellLeft, MSFlex.Top + MSFlex.CellTop + 3600, MSFlex.TextMatrix(nFila, 6), "Sustento:"
MSFlex.TextMatrix(nFila, 6) = frmLogProSelEspecificaciones.vpTexto
End Sub

Sub FormaFlex()
Dim i As Integer
MSFlex.Clear
MSFlex.Rows = 2
MSFlex.RowHeight(-1) = 260
MSFlex.RowHeight(0) = 320
MSFlex.RowHeight(1) = 8
MSFlex.ColWidth(0) = 0
MSFlex.ColWidth(1) = 0:         MSFlex.TextMatrix(0, 1) = ""
MSFlex.ColWidth(2) = 0:         MSFlex.TextMatrix(0, 2) = ""
MSFlex.ColWidth(3) = 3000:      MSFlex.TextMatrix(0, 3) = "Descripción"
MSFlex.ColWidth(4) = 1200:      MSFlex.TextMatrix(0, 4) = "Unidad":   MSFlex.ColAlignment(4) = 1
MSFlex.ColWidth(5) = 900: MSFlex.TextMatrix(0, 5) = "   Cantidad"
MSFlex.ColWidth(6) = 6000
MSFlex.ColWidth(7) = 0

MSFlexLista.Clear
MSFlexLista.Rows = 2
MSFlexLista.RowHeight(-1) = 260
MSFlexLista.RowHeight(0) = 320
MSFlexLista.RowHeight(1) = 8
MSFlexLista.ColWidth(0) = 0
MSFlexLista.ColWidth(1) = 300:       MSFlexLista.TextMatrix(0, 1) = ""
MSFlexLista.ColWidth(2) = 0:         MSFlexLista.TextMatrix(0, 2) = ""
MSFlexLista.ColWidth(3) = 4500:      MSFlexLista.TextMatrix(0, 3) = "Descripción"
MSFlexLista.ColWidth(4) = 1200:      MSFlexLista.TextMatrix(0, 4) = "Unidad":   MSFlex.ColAlignment(4) = 1
MSFlexLista.ColWidth(5) = 1000:      MSFlexLista.TextMatrix(0, 5) = "   Cantidad"
MSFlexLista.ColWidth(6) = 4000
MSFlexLista.ColWidth(7) = 0
End Sub


'Private Sub MSFlex_DblClick()
'On Error GoTo MSFlexErr
'    Dim i As Integer, bTipo As Boolean
'    With MSFlex
'        If Trim(.TextMatrix(.row, 1)) = "-" Then
'           .TextMatrix(.row, 1) = "+"
'           i = .row + 1
'           bTipo = True
'        ElseIf Trim(.TextMatrix(.row, 1)) = "+" Then
'           .TextMatrix(.row, 1) = "-"
'           i = .row + 1
'           bTipo = False
'        End If
'        Do While i < .Rows
'            If Trim(.TextMatrix(i, 1)) = "+" Or Trim(.TextMatrix(i, 1)) = "-" Or Trim(.TextMatrix(i, 0)) = "" Then
'                Exit Sub
'            End If
'
'            If bTipo Then
'                .RowHeight(i) = 0
'            Else
'                .RowHeight(i) = 260
'            End If
'            i = i + 1
'        Loop
'    End With
'Exit Sub
'MSFlexErr:
'    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
'End Sub

Function YaEsta(vBSCod As String) As Boolean
Dim i As Integer, n As Integer
YaEsta = False
n = MSFlex.Rows - 1

For i = 1 To n
    If MSFlex.TextMatrix(i, 1) = vBSCod Then
       YaEsta = True
       Exit Function
    End If
Next
End Function

'*********************************************************************
'PROCEDIMIENTOS DEL FLEX
'*********************************************************************
'Private Sub MSFlex_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyDelete And MSFlex.Col = 3 Then
'   MSFlex.TextMatrix(MSFlex.Row, 3) = ""
'End If
'End Sub

Private Sub MSFlex_KeyPress(KeyAscii As Integer)
If MSFlex.Col = 5 And nEditable Then
   EditaFlex MSFlex, txtEdit, KeyAscii
End If
End Sub

Sub EditaFlex(MSFlex As Control, Edt As Control, KeyAscii As Integer)
If InStr("0123456789", Chr(KeyAscii)) Then
   Select Case KeyAscii
    Case 0 To 32
         Edt = MSFlex
         Edt.SelStart = 1000
    Case Else
         Edt = Chr(KeyAscii)
         Edt.SelStart = 1
   End Select
   Edt.Move MSFlex.Left + MSFlex.CellLeft - 15, MSFlex.Top + MSFlex.CellTop - 15, _
            MSFlex.CellWidth, MSFlex.CellHeight
   Edt.Visible = True
   Edt.SetFocus
End If
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
nKeyAscii = KeyAscii
KeyAscii = DigNumEnt(KeyAscii)
If KeyAscii = Asc(vbCr) Then
   KeyAscii = 0
   txtEdit = FNumero(txtEdit)
End If
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
EditKeyCode MSFlex, txtEdit, KeyCode, Shift
End Sub

Sub EditKeyCode(MSFlex As Control, Edt As Control, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 27
         Edt.Visible = False
         MSFlex.SetFocus
    Case 13
         MSFlex.SetFocus
    Case 37                     'Izquierda
         MSFlex.SetFocus
         DoEvents
         If MSFlex.Col > 1 Then
            MSFlex.Col = MSFlex.Col - 1
         End If
    Case 39                     'Derecha
         MSFlex.SetFocus
         DoEvents
         If MSFlex.Col < MSFlex.Cols - 1 Then
            MSFlex.Col = MSFlex.Col + 1
         End If
    Case 38
         MSFlex.SetFocus
         DoEvents
         If MSFlex.row > MSFlex.FixedRows + 1 Then
            MSFlex.row = MSFlex.row - 1
         End If
    Case 40
         MSFlex.SetFocus
         DoEvents
         'If MSFlex.Row < MSFlex.FixedRows - 1 Then
         If MSFlex.row < MSFlex.Rows - 1 Then
            MSFlex.row = MSFlex.row + 1
         End If
End Select
End Sub

Private Sub MSFlex_RowColChange()
If Len(Trim(MSFlex.TextMatrix(MSFlex.row, 2))) > 0 Then
   cmdAgregarPers.Enabled = True
   cmdQuitarPers.Enabled = True
Else
   cmdAgregarPers.Enabled = False
   cmdQuitarPers.Enabled = False
End If
ListaPersonas MSFlex.TextMatrix(MSFlex.row, 2)
End Sub

Private Sub MSFlex_GotFocus()
If Len(Trim(MSFlex.TextMatrix(MSFlex.row, 2))) > 0 Then
   cmdAgregarPers.Enabled = True
   cmdQuitarPers.Enabled = True
Else
   cmdAgregarPers.Enabled = False
   cmdQuitarPers.Enabled = False
End If
ListaPersonas MSFlex.TextMatrix(MSFlex.row, 2)
If txtEdit.Visible = False Then Exit Sub
MSFlex = txtEdit
txtEdit.Visible = False
End Sub

Private Sub MSFlex_LeaveCell()
If txtEdit.Visible = False Then Exit Sub
MSFlex = txtEdit
txtEdit.Visible = False
End Sub


Sub GeneraMeses()
Dim i As Integer
For i = 1 To 12
    cboMes.AddItem Format("01/" & i & "/" & Year(gdFecSis), "mmmm")
Next
End Sub

Private Function VerificaPersona(ByVal pcPersCod As String, ByVal pcBSCod As String) As Boolean
Dim i As Integer, n As Integer
On Error GoTo VerificaPersonaErr
    
VerificaPersona = False
n = MSPers.Rows - 1
For i = 1 To n
    If MSPers.TextMatrix(i, 1) = pcPersCod And MSPers.TextMatrix(i, 2) = pcBSCod Then
       VerificaPersona = True
       Exit Function
    End If
Next
Exit Function

VerificaPersonaErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Function

Private Function VerificaCantidad(ByVal pnIndex As Integer, Optional TodaVerificacion As Boolean = False) As Boolean
Dim i As Integer, n As Integer, nCantPers As Integer, nCantidad As Integer
Dim cBSCod As String
On Error GoTo VerificaPersonaErr
    
VerificaCantidad = True
n = MSPers.Rows - 1
cBSCod = MSFlex.TextMatrix(pnIndex, 2)
nCantidad = Val(MSFlex.TextMatrix(pnIndex, 5))
nCantPers = 0
For i = 1 To n
    If MSPers.TextMatrix(i, 1) = cBSCod Then
       nCantPers = nCantPers + 1
    End If
Next

If TodaVerificacion Then
   If nCantPers = 0 Then
      MsgBox "Debe indicar al menos una persona para el requerimiento...", vbInformation, "Aviso"
      VerificaCantidad = False
      Exit Function
   End If

   'If nCantPers < nCantidad Then
   '   MsgBox "El número de personas debe coincidir con la cantidad requerida...", vbInformation, "Aviso"
   '   VerificaCantidad = False
   '   Exit Function
   'End If
End If

If Not TodaVerificacion Then
   If nCantPers + 1 > nCantidad Then
      MsgBox "Ya se ha completado el número de personas " & Space(10) & vbCrLf & "   de acuerdo a la cantidad del requerimiento..." + Space(10), vbInformation, "Confirme"
      VerificaCantidad = False
      Exit Function
   End If
End If

Exit Function

VerificaPersonaErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Function

Private Sub ConsultarRequerimientos(ByVal pcPersCod As String)
On Error GoTo ConsultarRequerimientosErr
    Dim oCon As DConecta, sSQL As String, Rs As ADODB.Recordset, nProselNro As Integer, i As Integer
    Set oCon = New DConecta
    sSQL = "select r.nProSelReqNro, b.cBSDescripcion, d.nCantidad, u.cUnidad, d.cSustento " & _
           " from LogProSelReq r " & _
           " inner join LogProSelReqDetalle d on r.nProSelReqNro = d.nProSelReqNro " & _
           " inner join LogProSelBienesServicios b on d.cBSCod = b.cProSelBSCod " & _
           " inner join (select nConsValor as nBSUnidad, cConsDescripcion as cUnidad from Constante where nConsCod = 9097) u on b.nBSUnidad = u.nBSUnidad " & _
           " where r.cPersCod = '" & pcPersCod & "' and nProSelNro = 0"
    If oCon.AbreConexion Then
        Set Rs = oCon.CargaRecordSet(sSQL)
        oCon.CierraConexion
    End If
    With MSFlexLista
        Do While Not Rs.EOF
            If nProselNro <> Rs!nProSelReqNro Then
                i = i + 1
                .Col = 1
                InsRow MSFlexLista, i
                .row = i
                .CellFontBold = True
                .CellFontSize = 10
                .CellAlignment = 4
                .TextMatrix(i, 1) = "+"
                .TextMatrix(i, 2) = Rs!nProSelReqNro
                .TextMatrix(i, 3) = Rs!cSustento
                nProselNro = Rs!nProSelReqNro
            End If
            i = i + 1
            InsRow MSFlexLista, i
            .RowHeight(i) = 0
            .TextMatrix(i, 3) = Rs!cBSDescripcion
            .TextMatrix(i, 4) = Rs!cUnidad
            .TextMatrix(i, 5) = Rs!nCantidad
            Rs.MoveNext
        Loop
    End With
    Exit Sub
ConsultarRequerimientosErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Sub

Private Function VerificarAreaAgencia(ByVal vPersCod As String) As Boolean
On Error GoTo VerificarAreaAgenciaErr
    Dim Rs As ADODB.Recordset, oCon As DConecta, sAge As String, sAre As String, sSQL As String
    Set oCon = New DConecta
    sSQL = "select * from rrhh where cperscod ='" & vPersCod & "'"
    If oCon.AbreConexion Then
      Set Rs = oCon.CargaRecordSet(sSQL)
      If Not Rs.EOF Then
      sAge = Rs!cAgenciaAsig
      sAre = Rs!cAreaCod
      If cRHAgeCod = sAge And cRHAreaCod = sAre Then
        VerificarAreaAgencia = True
      Else
        VerificarAreaAgencia = False
      End If
      End If
      oCon.CierraConexion
    End If
    Exit Function
VerificarAreaAgenciaErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Function

Private Sub cmdAgregarPers_Click()
Dim X As UPersona, i As Integer

If Not VerificaCantidad(MSFlex.row) Then
   Exit Sub
End If

Set X = frmBuscaPersona.Inicio(True)
If X Is Nothing Then
   Exit Sub
End If
    

   
    
If Len(Trim(X.sPersNombre)) > 0 Then
   If VerificaPersona(X.sPersCod, MSFlex.TextMatrix(MSFlex.row, 2)) Then
      MsgBox "Persona ya fue Ingresada", vbInformation, "Aviso"
      Exit Sub
   End If
   'If Not VerificaCantidad(Val(MSFlex.TextMatrix(MSFlex.row, 5)), MSFlex.TextMatrix(MSFlex.row, 2)) Then
   '   MsgBox "Nro de Personas debe Conincidir con la Cantidad", vbInformation, "Aviso"
   '   Exit Sub
   'End If
   i = MSPers.Rows - 1
   If i = 1 Then
      If Len(MSPers.TextMatrix(i, 2)) = 0 Then
         i = 0
      End If
   End If
   
   With MSFlex
        If Not VerificarAreaAgencia(X.sPersCod) Then
           MsgBox "Debe ser de la misma area y agencia", vbInformation, "Aviso"
           Exit Sub
        End If
        
        i = i + 1
        InsRow MSPers, i
        MSPers.TextMatrix(i, 1) = MSFlex.TextMatrix(MSFlex.row, 2)
        MSPers.TextMatrix(i, 2) = X.sPersCod
        MSPers.TextMatrix(i, 3) = X.sPersNombre
        ListaPersonas MSFlex.TextMatrix(MSFlex.row, 2)
        '    .Col = 1
        '    .CellFontBold = True
        '    .CellFontSize = 10
        '    .CellAlignment = 4
        '    .TextMatrix(i, 1) = "-"
        '    InsRow MSFlex, .Rows
        '    .TextMatrix(.Rows - 1, 0) = X.sPersCod
        '    .TextMatrix(.Rows - 1, 2) = .TextMatrix(i, 2)
        '    .TextMatrix(.Rows - 1, 3) = X.sPersNombre
        '    .TextMatrix(.Rows - 1, 7) = i
    End With
End If


End Sub

Private Sub cmdQuitarPers_Click()
Dim k As Integer

End Sub

Sub ListaPersonas(ByVal psBSCod As String)
Dim i As Integer, k As Integer, n As Integer
n = MSPers.Rows - 1
For i = 1 To n
    If MSPers.TextMatrix(i, 1) = psBSCod Then
       MSPers.RowHeight(i) = 270
    Else
       MSPers.RowHeight(i) = 0
    End If
Next
End Sub

Sub FlexPersonas()
MSPers.Clear
MSPers.Rows = 2
MSPers.RowHeight(0) = 8
MSPers.RowHeight(1) = 0
MSPers.ColWidth(0) = 0
MSPers.ColWidth(1) = 0
MSPers.ColWidth(2) = 1200
MSPers.ColWidth(3) = 6000
MSPers.ColWidth(4) = 0
MSPers.ColWidth(5) = 0
MSPers.ColWidth(6) = 0
End Sub
