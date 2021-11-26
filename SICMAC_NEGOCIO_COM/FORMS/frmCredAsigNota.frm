VERSION 5.00
Begin VB.Form frmCredAsigNota 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asignacion de Nota de Analista"
   ClientHeight    =   6780
   ClientLeft      =   2460
   ClientTop       =   1200
   ClientWidth     =   7200
   Icon            =   "frmCredAsigNota.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   135
      TabIndex        =   38
      Top             =   5895
      Width           =   6975
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "&Grabar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   255
         TabIndex        =   41
         Top             =   225
         Width           =   1305
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5580
         TabIndex        =   40
         Top             =   210
         Width           =   1305
      End
      Begin VB.CommandButton CmdNuevo 
         Caption         =   "&Nuevo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   4185
         TabIndex        =   39
         Top             =   225
         Width           =   1305
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Credito"
      Height          =   1245
      Left            =   105
      TabIndex        =   31
      Top             =   0
      Width           =   6990
      Begin VB.Frame FraListaCred 
         Caption         =   "&Lista Creditos"
         Height          =   960
         Left            =   5025
         TabIndex        =   34
         Top             =   180
         Width           =   1875
         Begin VB.ListBox LstCred 
            Height          =   645
            ItemData        =   "frmCredAsigNota.frx":030A
            Left            =   75
            List            =   "frmCredAsigNota.frx":030C
            TabIndex        =   35
            Top             =   225
            Width           =   1725
         End
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4035
         TabIndex        =   33
         Top             =   525
         Width           =   900
      End
      Begin SICMACT.ActXCodCta ActxCta 
         Height          =   435
         Left            =   195
         TabIndex        =   32
         Top             =   495
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   767
         Texto           =   "Credito :"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
   End
   Begin VB.Frame FraNota2 
      Enabled         =   0   'False
      Height          =   1845
      Left            =   135
      TabIndex        =   22
      Top             =   4050
      Width           =   6975
      Begin VB.Frame FraNota 
         Caption         =   "Nota"
         Height          =   1560
         Left            =   90
         TabIndex        =   25
         ToolTipText     =   "Seleccione una de las opciones para la asignación de la Nota"
         Top             =   150
         Width           =   1890
         Begin VB.OptionButton OptNota2 
            Caption         =   "Dias atraso  = 0"
            Height          =   270
            Index           =   0
            Left            =   135
            TabIndex        =   30
            ToolTipText     =   "Seleccione una de las opciones para la asignación de la Nota"
            Top             =   210
            Width           =   1650
         End
         Begin VB.OptionButton OptNota2 
            Caption         =   "Dias atraso  <= 3"
            Height          =   270
            Index           =   1
            Left            =   135
            TabIndex        =   29
            ToolTipText     =   "Seleccione una de las opciones para la asignación de la Nota"
            Top             =   450
            Width           =   1650
         End
         Begin VB.OptionButton OptNota2 
            Caption         =   "Dias atraso  <= 5"
            Height          =   270
            Index           =   2
            Left            =   135
            TabIndex        =   28
            ToolTipText     =   "Seleccione una de las opciones para la asignación de la Nota"
            Top             =   705
            Width           =   1650
         End
         Begin VB.OptionButton OptNota2 
            Caption         =   "Dias atraso  <= 7"
            Height          =   270
            Index           =   3
            Left            =   135
            TabIndex        =   27
            ToolTipText     =   "Seleccione una de las opciones para la asignación de la Nota"
            Top             =   945
            Width           =   1650
         End
         Begin VB.OptionButton OptNota2 
            Caption         =   "Dias atraso  > 7"
            Height          =   270
            Index           =   4
            Left            =   135
            TabIndex        =   26
            Top             =   1215
            Width           =   1650
         End
      End
      Begin VB.Frame fraComentario 
         Caption         =   "Comentario "
         Height          =   1545
         Left            =   2010
         TabIndex        =   23
         Top             =   150
         Width           =   4815
         Begin VB.TextBox TxtComenta 
            Height          =   1170
            Left            =   90
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   24
            Top             =   240
            Width           =   4605
         End
      End
   End
   Begin VB.Frame FraCliente 
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   105
      TabIndex        =   14
      Top             =   1320
      Width           =   7020
      Begin VB.Label Label12 
         Caption         =   "Datos del Cliente"
         Height          =   195
         Left            =   225
         TabIndex        =   21
         Top             =   -30
         Width           =   1275
      End
      Begin VB.Label LblNomCliente 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   1080
         TabIndex        =   20
         Top             =   600
         Width           =   5820
      End
      Begin VB.Label Label3 
         Caption         =   "Doc. Identidad:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4125
         TabIndex        =   19
         Top             =   255
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Código:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   600
         Width           =   735
      End
      Begin VB.Label LblCodCliente 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   1080
         TabIndex        =   16
         Top             =   240
         Width           =   1470
      End
      Begin VB.Label LblDIdent 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   5310
         TabIndex        =   15
         Top             =   255
         Width           =   1575
      End
   End
   Begin VB.Frame FraCredito 
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   1680
      Left            =   120
      TabIndex        =   0
      Top             =   2325
      Width           =   6990
      Begin VB.Label Label10 
         Caption         =   "Nota Sistema :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5025
         TabIndex        =   43
         Top             =   390
         Width           =   1290
      End
      Begin VB.Label LblNotaSist 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6390
         TabIndex        =   42
         Top             =   375
         Width           =   420
      End
      Begin VB.Label LblDestinoC 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   990
         TabIndex        =   37
         Top             =   915
         Width           =   1770
      End
      Begin VB.Label LblMontoS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   990
         TabIndex        =   36
         Top             =   1275
         Width           =   1770
      End
      Begin VB.Label Label11 
         Caption         =   "Datos del Crédito"
         Height          =   255
         Left            =   225
         TabIndex        =   13
         Top             =   -30
         Width           =   1260
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo"
         Height          =   255
         Left            =   105
         TabIndex        =   12
         Top             =   255
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Prestamo :"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1290
         Width           =   840
      End
      Begin VB.Label Label7 
         Caption         =   "Moneda:"
         Height          =   255
         Left            =   3270
         TabIndex        =   10
         Top             =   1275
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Destino"
         Height          =   255
         Left            =   135
         TabIndex        =   9
         Top             =   975
         Width           =   735
      End
      Begin VB.Label LblTipoC 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   990
         TabIndex        =   8
         Top             =   255
         Width           =   3885
      End
      Begin VB.Label LblMoneda 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   4005
         TabIndex        =   7
         Top             =   1290
         Width           =   900
      End
      Begin VB.Label Label16 
         Caption         =   "Dias Atraso Acum."
         Height          =   225
         Left            =   5025
         TabIndex        =   6
         Top             =   975
         Width           =   1365
      End
      Begin VB.Label LblDiasAcum 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   6405
         TabIndex        =   5
         Top             =   960
         Width           =   420
      End
      Begin VB.Label LblNota2 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   6405
         TabIndex        =   4
         Top             =   1335
         Width           =   420
      End
      Begin VB.Label Label15 
         Caption         =   "Nota Asignada"
         Height          =   225
         Left            =   5025
         TabIndex        =   3
         Top             =   1305
         Width           =   1290
      End
      Begin VB.Label Label9 
         Caption         =   "Analista:"
         Height          =   255
         Left            =   135
         TabIndex        =   2
         Top             =   630
         Width           =   720
      End
      Begin VB.Label LblAnalista 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   990
         TabIndex        =   1
         Top             =   600
         Width           =   3900
      End
   End
End
Attribute VB_Name = "frmCredAsigNota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Function CargaDatos(ByVal psCtaCod As String) As Boolean
Dim oDCred As COMDCredito.DCOMCredito
Dim R As ADODB.Recordset
Dim nInd As Integer
    On Error GoTo ErrorCargaDatos
    Set oDCred = New COMDCredito.DCOMCredito
    Set R = oDCred.RecuperaDatosComunes(psCtaCod, False)
    Set oDCred = Nothing
    If Not R.BOF And Not R.EOF Then
        CargaDatos = True
        LblCodCliente.Caption = R!cPersCod
        LblDIdent.Caption = IIf(IsNull(R!DNI), "", R!DNI)
        LblNomCliente.Caption = PstaNombre(R!cTitular)
        LblTipoC.Caption = Trim(R!cTipoCredDescrip)
        LblDestinoC.Caption = Trim(R!cDestinoDescripcion)
        LblAnalista.Caption = PstaNombre(R!cAnalista)
        LblDiasAcum.Caption = IIf(IsNull(R!nDiasAtrasoAcum), 0, R!nDiasAtrasoAcum)
        LblMontoS.Caption = Format(R!nMontoCol, "#0.00")
        LblMoneda.Caption = Trim(R!cMoneda)
        nInd = IIf(IsNull(R!nNota), 1, R!nNota)
        If nInd > 0 Then OptNota2(nInd - 1).value = True
        LblNota2.Caption = IIf(IsNull(R!nNota), "0", R!nNota)
        LblNotaSist.Caption = IIf(IsNull(R!nNotaSist), "", R!nNotaSist)
    Else
        CargaDatos = False
    End If
    R.Close
    Set R = Nothing
    Exit Function

ErrorCargaDatos:
        MsgBox Err.Description, vbCritical, "Aviso"

End Function

Private Sub LimpiaPantalla()
    LimpiaControles Me, True
    ActxCta.NroCuenta = ""
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    OptNota2(0).value = True
End Sub

Private Sub HabilitaActualizacion(ByVal pbHabilita As Boolean)
    FraNota2.Enabled = pbHabilita
    CmdGrabar.Enabled = pbHabilita
    Frame3.Enabled = Not pbHabilita
End Sub
Private Sub ActxCta_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorActxCta_KeyPress
    If KeyAscii = 13 Then
        If Not CargaDatos(ActxCta.NroCuenta) Then
            HabilitaActualizacion False
            MsgBox "No se pudo encontrar el Credito, o el Credito No esta Vigente", vbInformation, "Aviso"
        Else
            HabilitaActualizacion True
        End If
    End If
    Exit Sub

ErrorActxCta_KeyPress:
        MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub cmdGrabar_Click()
Dim oNCred As COMNCredito.NCOMCredito
    If MsgBox("Se va a Actualizar la Nota del Credito, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    HabilitaActualizacion False
    Set oNCred = New COMNCredito.NCOMCredito
    Call oNCred.ActualizarNotaCredito(ActxCta.NroCuenta, CInt(LblNota2.Caption), gdFecSis, TxtComenta.Text)
    Set oNCred = Nothing
    Call cmdNuevo_Click
End Sub

Private Sub cmdNuevo_Click()
    Call LimpiaPantalla
    HabilitaActualizacion False
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 And ActxCta.Enabled = True Then 'F12
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.Inicia(gColPYMEEmp, False)
        If sCuenta <> "" Then
            ActxCta.NroCuenta = sCuenta
            ActxCta.SetFocusCuenta
        End If
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    CentraForm Me
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
End Sub

Private Sub cmdBuscar_Click()
Dim oCredito As COMDCredito.DCOMCredito
Dim R As ADODB.Recordset
Dim oPers As comdpersona.UCOMPersona

    On Error GoTo ErrorCmdBuscar_Click
    LstCred.Clear
    Set oPers = frmBuscaPersona.Inicio()
    If Not oPers Is Nothing Then
        Set oCredito = New COMDCredito.DCOMCredito
        Set R = oCredito.RecuperaCreditosVigentes(oPers.sPersCod, , Array(gColocEstVigMor, gColocEstVigNorm, gColocEstVigVenc))
        Do While Not R.EOF
            LstCred.AddItem R!cCtaCod
            R.MoveNext
        Loop
        R.Close
        Set R = Nothing
        Set oCredito = Nothing
    End If
    If LstCred.ListCount = 0 Then
        MsgBox "El Cliente No Tiene Creditos Vigentes", vbInformation, "Aviso"
    End If
    
    Exit Sub

ErrorCmdBuscar_Click:
    MsgBox Err.Description, vbCritical, "Aviso"

End Sub

Private Sub LstCred_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If LstCred.ListCount > 0 And LstCred.ListIndex <> -1 Then
            ActxCta.NroCuenta = LstCred.Text
            ActxCta.SetFocus
        End If
    End If
End Sub

Private Sub OptNota2_Click(Index As Integer)
    LblNota2.Caption = Index + 1
End Sub
