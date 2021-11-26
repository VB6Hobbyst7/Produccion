VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{9E7F5C4C-A058-11D4-8E95-444553540000}#28.0#0"; "TBVariable.ocx"
Begin VB.Form frmPosSBS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Central del Sistema Financiero ... SICMAC-I"
   ClientHeight    =   6390
   ClientLeft      =   1440
   ClientTop       =   3105
   ClientWidth     =   10335
   Icon            =   "frmPosSBS.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   10335
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   6075
      TabIndex        =   39
      Top             =   5940
      Width           =   1380
   End
   Begin VB.Frame Frame4 
      Caption         =   " Busqueda "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   180
      TabIndex        =   11
      Top             =   450
      Width           =   10050
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   360
         Left            =   6450
         TabIndex        =   22
         Top             =   300
         Width           =   1335
      End
      Begin VB.TextBox txtDato 
         Height          =   330
         Left            =   4950
         TabIndex        =   21
         Top             =   705
         Width           =   4400
      End
      Begin VB.Frame Frame3 
         Caption         =   "Busqueda Por..."
         Height          =   855
         Left            =   3240
         TabIndex        =   17
         Top             =   225
         Width           =   1695
         Begin VB.OptionButton optBuscar 
            Caption         =   "Codi&go"
            Height          =   330
            Index           =   2
            Left            =   90
            TabIndex        =   20
            Top             =   930
            Width           =   1440
         End
         Begin VB.OptionButton optBuscar 
            Caption         =   "&Documento"
            Height          =   315
            Index           =   1
            Left            =   90
            TabIndex        =   19
            Top             =   480
            Width           =   1515
         End
         Begin VB.OptionButton optBuscar 
            Caption         =   "Apellidos/&Razon"
            Height          =   330
            Index           =   0
            Left            =   90
            TabIndex        =   18
            Top             =   180
            Value           =   -1  'True
            Width           =   1470
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " Selector "
         Height          =   855
         Left            =   165
         TabIndex        =   12
         Top             =   225
         Width           =   3060
         Begin VB.ComboBox cboYeaPro 
            CausesValidation=   0   'False
            Height          =   315
            Left            =   60
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   405
            Width           =   1440
         End
         Begin VB.ComboBox cboMesYea 
            CausesValidation=   0   'False
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   390
            Width           =   1440
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Año a Procesar"
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
            Left            =   105
            TabIndex        =   16
            Top             =   225
            Width           =   1320
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Mes a Procesar"
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
            Left            =   1605
            TabIndex        =   15
            Top             =   210
            Width           =   1335
         End
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Dato a Buscar:"
         Height          =   195
         Left            =   4965
         TabIndex        =   23
         Top             =   360
         Width           =   1065
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   " Información Detallada "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2820
      Left            =   105
      TabIndex        =   8
      Top             =   3060
      Width           =   10065
      Begin MSComctlLib.ListView lvwInfDet 
         CausesValidation=   0   'False
         Height          =   2460
         Left            =   45
         TabIndex        =   9
         Top             =   225
         Width           =   9945
         _ExtentX        =   17542
         _ExtentY        =   4339
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo IFI"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Institucion Financiera"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Tipo Credito"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Rubros"
            Object.Width           =   6703
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Atrazo"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Saldo"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Calificacion"
            Object.Width           =   1411
         EndProperty
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   345
      Left            =   8820
      TabIndex        =   6
      Top             =   5940
      Width           =   1365
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   90
      TabIndex        =   3
      Top             =   1635
      Width           =   10050
      Begin VB.TextBox txtEmpInf 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2055
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   915
         Width           =   570
      End
      Begin VB.TextBox txtPerson 
         CausesValidation=   0   'False
         Height          =   330
         Left            =   7350
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   225
         Width           =   2610
      End
      Begin TBVariable.ctlTBVar txtCalif0 
         CausesValidation=   0   'False
         Height          =   270
         Left            =   6075
         TabIndex        =   25
         Top             =   915
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   476
         Text            =   "0.00"
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         ForeColor       =   0
         Locked          =   -1  'True
         CharAccept      =   1
      End
      Begin VB.TextBox txtCodSBS 
         CausesValidation=   0   'False
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
         Height          =   330
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   570
         Width           =   1665
      End
      Begin VB.TextBox txtDocIde 
         CausesValidation=   0   'False
         Height          =   345
         Left            =   3570
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   570
         Width           =   1485
      End
      Begin VB.TextBox txtNomcli 
         CausesValidation=   0   'False
         Height          =   330
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   225
         Width           =   6330
      End
      Begin TBVariable.ctlTBVar txtCalif1 
         CausesValidation=   0   'False
         Height          =   270
         Left            =   6855
         TabIndex        =   27
         Top             =   915
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   476
         Text            =   "0.00"
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         ForeColor       =   0
         Locked          =   -1  'True
         CharAccept      =   1
      End
      Begin TBVariable.ctlTBVar txtCalif2 
         CausesValidation=   0   'False
         Height          =   270
         Left            =   7635
         TabIndex        =   29
         Top             =   915
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   476
         Text            =   "0.00"
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         ForeColor       =   0
         Locked          =   -1  'True
         CharAccept      =   1
      End
      Begin TBVariable.ctlTBVar txtCalif3 
         CausesValidation=   0   'False
         Height          =   270
         Left            =   8415
         TabIndex        =   31
         Top             =   915
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   476
         Text            =   "0.00"
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         ForeColor       =   0
         Locked          =   -1  'True
         CharAccept      =   1
      End
      Begin TBVariable.ctlTBVar txtCalif4 
         CausesValidation=   0   'False
         Height          =   270
         Left            =   9195
         TabIndex        =   33
         Top             =   915
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   476
         Text            =   "0.00"
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         ForeColor       =   0
         Locked          =   -1  'True
         CharAccept      =   1
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Empresas Informantes"
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
         Left            =   105
         TabIndex        =   36
         Top             =   1005
         Width           =   1875
      End
      Begin VB.Label Label11 
         BackColor       =   &H00808000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Calif. 4 "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   9195
         TabIndex        =   34
         Top             =   675
         Width           =   765
      End
      Begin VB.Label Label10 
         BackColor       =   &H00808000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Calif. 3 "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   8415
         TabIndex        =   32
         Top             =   675
         Width           =   765
      End
      Begin VB.Label Label8 
         BackColor       =   &H00808000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Calif. 2 "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   7635
         TabIndex        =   30
         Top             =   675
         Width           =   765
      End
      Begin VB.Label Label7 
         BackColor       =   &H00808000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Calif. 1 "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   6855
         TabIndex        =   28
         Top             =   675
         Width           =   765
      End
      Begin VB.Label Label4 
         BackColor       =   &H00808000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Calif. 0 "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   6075
         TabIndex        =   26
         Top             =   675
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Doc. Ide :"
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
         Left            =   2805
         TabIndex        =   10
         Top             =   675
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nombre :"
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
         Left            =   150
         TabIndex        =   5
         Top             =   285
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cod.SBS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   150
         TabIndex        =   4
         Top             =   645
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   345
      Left            =   7455
      TabIndex        =   7
      Top             =   5940
      Width           =   1365
   End
   Begin VB.Label lblTitulo 
      AutoSize        =   -1  'True
      Caption         =   "INFORMACION SBS A: JUNIO - 2004"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   285
      TabIndex        =   38
      Top             =   105
      Width           =   4530
   End
   Begin VB.Label lblRutAge 
      BorderStyle     =   1  'Fixed Single
      Height          =   75
      Left            =   -45
      TabIndex        =   24
      Top             =   -15
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmPosSBS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sRutFil As String
Private nTipBus As Integer
Private bPriVez As Boolean
Private gsBaseRCC As String

Private Function VerificaRuta(ByVal psYeaPro As String, ByVal psMesPro As String) As Boolean
Dim SQL As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta
Set oCon = New DConecta

VerificaRuta = False

oCon.AbreConexion
SQL = "SELECT month(MAX(FEC_REP)) as nMes, year(MAX(FEC_REP)) as nAnio FROM " & gsBaseRCC & "RCCTOTAL"
Set rs = oCon.CargaRecordSet(SQL)
If Not rs.EOF And Not rs.BOF Then
    If Val(rs!nAnio) = Val(psYeaPro) And Val(rs!nMes) = Val(psMesPro) Then
        VerificaRuta = True
    End If
End If
rs.Close
Set rs = Nothing
oCon.CierraConexion
Set oCon = Nothing
End Function

Private Function GetValidaFecha(ByVal psYeaPro As String, ByVal psMesPro As String) As Boolean
Dim SQL As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta
Set oCon = New DConecta

GetValidaFecha = False

oCon.AbreConexion
SQL = "SELECT TOP 1 FEC_REP FROM " & gsBaseRCC & "RCCTOTAL WHERE month(FEC_REP) = " & psMesPro & " and  year(FEC_REP) =" & psYeaPro & ""
Set rs = oCon.CargaRecordSet(SQL)
If Not rs.EOF And Not rs.BOF Then
    GetValidaFecha = True
End If
rs.Close
Set rs = Nothing
oCon.CierraConexion
Set oCon = Nothing
End Function


Private Sub cboMesYea_Click()
    If bPriVez Then
       If GetValidaFecha(cboYeaPro, Right(Trim(cboMesYea), 2)) = False Then
          MsgBox "Informacion RCC no encontrada por la fecha seleccionada", vbInformation, "AVISO"
          Exit Sub
       Else
            txtNomcli.Text = ""
            txtPerson.Text = ""
            txtCodSBS.Text = ""
            txtDocIde.Text = ""
            txtEmpInf.Text = ""
            txtCalif0.Text = "0.00"
            txtCalif1.Text = "0.00"
            txtCalif2.Text = "0.00"
            txtCalif3.Text = "0.00"
            txtCalif4.Text = "0.00"
            lvwInfDet.ListItems.Clear
       End If
    End If
End Sub

Private Sub cmdBuscar_Click()
    Dim sCadSQl As String
    Dim rsDatCli As ADODB.Recordset
    Dim lsCodSBS As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
On Error GoTo ErrorBuscar

    If (GetValidaFecha(Trim(cboYeaPro.Text), Right(cboMesYea.Text, 2))) = False Then
       MsgBox "No existe Informacion del RCC al Período seleccionado"
       Exit Sub
    End If
    
    If Len(Trim(txtDato.Text)) <= 4 Then
        MsgBox "Debe Ingresar para la busqueda como minimo 4 caracteres"
        txtDato.SetFocus
        Exit Sub
    End If
    Set rsDatCli = New ADODB.Recordset
    
    oCon.AbreConexion
    
    lvwInfDet.ListItems.Clear
    If nTipBus = 0 Then
        frmMuestraPersRCC.Inicio gsBaseRCC, Trim(txtDato.Text), Right(cboMesYea, 2), cboYeaPro
        lsCodSBS = frmMuestraPersRCC.lsCodSBS
        txtDato.Text = Trim(frmMuestraPersRCC.lsNomCli)
        If lsCodSBS = "" Then
            MsgBox ("Por favor seleccione alguna persona")
            Exit Sub
        End If
        Set frmMuestraPersRCC = Nothing
        
        sCadSQl = "SELECT  Nom_Deu AS cNomCli,Tip_Pers AS cTipPers,Cod_Sbs AS CCODSBS, Cod_Doc_Id AS cNuDocI, Cod_Doc_Trib AS cNudOtr, Can_Ents AS nCanEmp, " _
            & "         Calif_0 AS nCalif0,Calif_1 AS nCalif1,Calif_2 AS nCalif2,Calif_3 AS nCalif3,Calif_4 AS nCalif4  " _
            & "FROM   " & gsBaseRCC & "RCCTOTAL  " _
            & "WHERE Cod_Sbs = '" & lsCodSBS & "' and month(Fec_Rep)=" & Val(Right(cboMesYea, 2)) & " and year(Fec_Rep)=" & Val(cboYeaPro)
    Else
        sCadSQl = "SELECT  Nom_Deu AS cNomCli,Tip_Pers AS cTipPers,Cod_Sbs AS CCODSBS, Cod_Doc_Id AS cNuDocI, Cod_Doc_Trib AS cNudOtr, Can_Ents AS nCanEmp, " _
                & "         Calif_0 AS nCalif0,Calif_1 AS nCalif1,Calif_2 AS nCalif2,Calif_3 AS nCalif3,Calif_4 AS nCalif4  " _
                & "FROM   " & gsBaseRCC & "RCCTOTAL " _
                & "WHERE (Cod_Doc_Id = '" & txtDato.Text & "' OR  Cod_Doc_Trib = '" & txtDato.Text & "') and month(Fec_Rep)=" & Val(Right(cboMesYea, 2)) & " and year(Fec_Rep)=" & Val(cboYeaPro)
    End If
    
    
    Set rsDatCli = oCon.CargaRecordSet(sCadSQl)
    If Not RSVacio(rsDatCli) Then
       With rsDatCli
           txtNomcli.Text = Trim(!cNomCli)
           txtPerson.Text = IIf(!cTipPers = "1", "NATURAL", IIf(!cTipPers = "2", "JURIDICA", ""))
           txtCodSBS.Text = !cCodSBS
           txtDocIde.Text = IIf(!cTipPers = "1", !cNuDocI, IIf(!cTipPers = "2", !cNudOtr, ""))
           txtEmpInf.Text = Format(!nCanEmp, "00")
           txtCalif0.Text = Format(!nCalif0, "#0.00")
           txtCalif1.Text = Format(!nCalif1, "#0.00")
           txtCalif2.Text = Format(!nCalif2, "#0.00")
           txtCalif3.Text = Format(!nCalif3, "#0.00")
           txtCalif4.Text = Format(!nCalif4, "#0.00")
       End With
       
       txtNomcli.Refresh
       txtPerson.Refresh
       txtCodSBS.Refresh
       txtDocIde.Refresh
       txtEmpInf.Refresh
       
            
       RecuperaCreditos
            
       
    Else
       MsgBox "Persona no Posee Registro en el Sistema Financiero..."
       txtNomcli.Text = ""
       txtPerson.Text = ""
       txtCodSBS.Text = ""
       txtDocIde.Text = ""
       txtEmpInf.Text = ""
       txtCalif0.Text = ""
       txtCalif1.Text = ""
       txtCalif2.Text = ""
       txtCalif3.Text = ""
       txtCalif4.Text = ""
       lvwInfDet.ListItems.Clear
      
    End If
    rsDatCli.Close
    Set rsDatCli = Nothing
    
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Sub
ErrorBuscar:
    Screen.MousePointer = 0
    
    MsgBox "Error N° " & Err.Number & " " & Err.Description, vbInformation, "Aviso"
    Exit Sub
End Sub

Private Sub cmdCancelar_Click()
txtNomcli.Text = ""
txtPerson.Text = ""
txtCodSBS.Text = ""
txtDocIde.Text = ""
txtEmpInf.Text = ""
txtCalif0.Text = ""
txtCalif1.Text = ""
txtCalif2.Text = ""
txtCalif3.Text = ""
txtCalif4.Text = ""
txtDato.Text = ""
lvwInfDet.ListItems.Clear

End Sub

Private Sub cmdImprimir_Click()
Dim lsCadImp  As String
Dim i As Integer
Dim lwItem As ListItem
lsCadImp = ""

lsCadImp = lsCadImp + CabeRepo(gsNomCmac, gsNomAge, "", "", Now(), PrnSet("B+") + "CONSULTA INFORMACION DEL SISTEMA FINANCIERO" + PrnSet("B-"), "INFORMACION SBS A :" & Trim(Left(cboMesYea, 50)) & " - " & Trim(cboYeaPro), "PERSONA :" & Trim(txtNomcli), "COD.SBS : " & Trim(txtCodSBS) + " - DOC.IDENT:" + Trim(txtDocIde), "0", 60)
Linea lsCadImp, "", 1
Linea lsCadImp, "NUMERO DE EMPRESAS : " & txtEmpInf
Linea lsCadImp, PrnSet("B+") & "CALIFICACION [0]: " & txtCalif0.Text & "% [1]: " & txtCalif1.Text & "% [2]: " & txtCalif2.Text & "%  [3]: " & txtCalif3.Text & "%  [4]: " & txtCalif4.Text & "%" & PrnSet("B-")

Linea lsCadImp, PrnSet("C+") + String(150, "-")
Linea lsCadImp, PrnSet("B+") + ImpreFormat("INSTITUCION FINANCIERA", 45) & ImpreFormat("TIPO DE CREDITO", 15) & ImpreFormat("RUBRO", 45) & ImpreFormat("DIAS.ATRAZO", 15) & ImpreFormat("SALDO", 6) & ImpreFormat("CALIF", 5) & PrnSet("B-")
Linea lsCadImp, String(150, "-")
For i = 1 To lvwInfDet.ListItems.Count
    Linea lsCadImp, PrnSet("C+") + ImpreFormat(Trim(lvwInfDet.ListItems(i).SubItems(1)), 45) & _
                     ImpreFormat(Trim(lvwInfDet.ListItems(i).SubItems(2)), 15) & ImpreFormat(Trim(lvwInfDet.ListItems(i).SubItems(3)), 40) & ImpreFormat(Val(lvwInfDet.ListItems(i).SubItems(4)), 12, 0) & ImpreFormat(Val(lvwInfDet.ListItems(i).SubItems(5)), 15, 2) & ImpreFormat(Val(lvwInfDet.ListItems(i).SubItems(6)), 4, 0)
Next
Linea lsCadImp, String(150, "-") + PrnSet("C-")

Linea lsCadImp, "USUARIO : " & gsCodUser

EnviaPrevio lsCadImp, "Consulta Central Sistema Financiero", 66
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim oCon As DConecta
Dim SQL As String
Dim rs As ADODB.Recordset
Set oCon = New DConecta
    CentraForm Me
    
gsBaseRCC = ""
oCon.AbreConexion
SQL = "SELECT nConsSisValor FROM CONSTSISTEMA where nConsSisCod= 144"
Set rs = oCon.CargaRecordSet(SQL)
If Not rs.EOF And Not rs.BOF Then
    gsBaseRCC = Trim(rs!nConsSisValor)
End If
rs.Close
Set rs = Nothing

oCon.CierraConexion
Set oCon = Nothing

CargarCombos
bPriVez = True
    
End Sub

Private Sub CargarCombos()
    Dim nYeaPro As Integer, M As Integer
    Dim nMesPro As Integer, nYeaIni As Integer, nYeaFin As Integer
    Dim i As Integer
    
    nYeaIni = 2000
    nYeaFin = Year(Now())
    nMesPro = Month(Now())
    
    For M = nYeaIni To nYeaFin
        cboYeaPro.AddItem Trim(CStr(M))
    Next
    cboYeaPro.ListIndex = (nYeaFin - nYeaIni)
    
    bPriVez = False
    cboMesYea.AddItem "ENERO" & Space(100) & "01"
    cboMesYea.AddItem "FEBRERO" & Space(100) & "02"
    cboMesYea.AddItem "MARZO" & Space(100) & "03"
    cboMesYea.AddItem "ABRIL" & Space(100) & "04"
    cboMesYea.AddItem "MAYO" & Space(100) & "05"
    cboMesYea.AddItem "JUNIO" & Space(100) & "06"
    cboMesYea.AddItem "JULIO" & Space(100) & "07"
    cboMesYea.AddItem "AGOSTO" & Space(100) & "08"
    cboMesYea.AddItem "SETIEMBRE" & Space(100) & "09"
    cboMesYea.AddItem "OCTUBRE" & Space(100) & "10"
    cboMesYea.AddItem "NOVIEMBRE" & Space(100) & "11"
    cboMesYea.AddItem "DICIEMBRE" & Space(100) & "12"
    
    cboMesYea.ListIndex = (nMesPro - 1)
    
    For i = cboMesYea.ListCount - 1 To 0 Step -1
        cboMesYea.ListIndex = i
        If (VerificaRuta(Trim(Str(nYeaFin)), Right(cboMesYea.Text, 2))) = True Then
            Me.lblTitulo = "INFORMACION SBS A: " & Trim(cboMesYea.Text) & " - " & Trim(Str(nYeaFin))
            Exit Sub
        Else
            Me.lblTitulo = ""
        End If
    Next
    
    bPriVez = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CierraConexion
End Sub

Private Sub optBuscar_Click(Index As Integer)
    nTipBus = Index
    Select Case Index
        Case 0
            txtDato.Width = 4400
            txtDato.MaxLength = 0
        Case 1
            txtDato.MaxLength = 25
            txtDato.Width = 2000
'        Case Else
'            txtDato.MaxLength = 12
'            txtDato.Width = 2000
    End Select
End Sub

Private Sub txtDato_KeyPress(KeyAscii As Integer)
    If optBuscar(0).value Then
       KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Else
       KeyAscii = NumerosEnteros(KeyAscii)
    End If
    If Len(Trim(cboYeaPro.Text)) = 0 Then
       MsgBox "Seleccione Año de Busqueda", vbInformation, "Aviso"
       cboYeaPro.SetFocus
       Exit Sub
    End If
    If Len(Trim(cboMesYea.Text)) = 0 Then
       MsgBox "Seleccione Mes de Busqueda", vbInformation, "Aviso"
       cboMesYea.SetFocus
       Exit Sub
    End If
    If KeyAscii = 13 Then
       If Len(Trim(txtDato)) = 0 Then
          MsgBox "Ingrese Dato a Buscar", vbInformation, "Aviso"
          txtDato.SetFocus
          Exit Sub
       End If
       cmdBuscar_Click
    End If
End Sub

Private Sub RecuperaCreditos()
    Dim sCadSQl As String
    Dim rsDatCli As ADODB.Recordset
    Dim lvwIteRec As ListItem
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    Dim lsTipCred As String
    
    oCon.AbreConexion
    
    Set rsDatCli = New ADODB.Recordset
    
    lvwInfDet.ListItems.Clear
    
    sCadSQl = "SELECT  Cod_Emp as cCodIfi, Tip_Credito as cTipCred, Cod_Cuenta cCtaCont, Condicion as nDiasAtr, Val_Saldo as nMonto, Clasificacion as cCalif " _
            & "FROM   " & gsBaseRCC & "RCCTOTALdet  " _
            & "WHERE    (Cod_Sbs = '" & Trim(txtCodSBS.Text) & "') " _
            & "         and month(dFecha)=" & Val(Right(cboMesYea, 2)) & " and year(dFecha)=" & Val(cboYeaPro)
    
    Set rsDatCli = oCon.CargaRecordSet(sCadSQl)
    If Not RSVacio(rsDatCli) Then
       With rsDatCli
           While Not .EOF
                Select Case !cTipCred
                    Case Is = "1"
                        lsTipCred = "COMERCIALES"
                    Case Is = "2"
                        lsTipCred = "MICROEMPRESA"
                    Case Is = "3"
                        lsTipCred = "CONSUMO"
                    Case Is = "4"
                        lsTipCred = "HIPOTECARIO"
                    Case Else
                        lsTipCred = "NO DEFINIDO"
                End Select
                
               Set lvwIteRec = lvwInfDet.ListItems.Add(, , !cCodIfi)
                   lvwIteRec.SubItems(1) = BuscaIFI(!cCodIfi)
                   lvwIteRec.SubItems(2) = lsTipCred '!cTipCred
                   lvwIteRec.SubItems(3) = BuscaCta(!cCtaCont)
                   lvwIteRec.SubItems(4) = !nDiasAtr
                   lvwIteRec.SubItems(5) = Format(!nMonto, "#0.00")
                   lvwIteRec.SubItems(6) = !cCalif
               .MoveNext
           Wend
       End With
    End If
    rsDatCli.Close
    Set rsDatCli = Nothing
    oCon.CierraConexion
    Set oCon = Nothing
End Sub

Private Function BuscaCta(ByVal psCodCta As String) As String
    Dim sCadSQl As String, sRutCta As String
    Dim rsDatCli As ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
    
    sCadSQl = "SELECT  Descripcio  " _
            & "FROM   " & gsBaseRCC & "RCDCuentas  " _
            & "WHERE (CuentaRCD = '" & Trim(psCodCta) & "')  ORDER BY CuentaRCD "
    
    Set rsDatCli = oCon.CargaRecordSet(sCadSQl)
    If Not RSVacio(rsDatCli) Then
       With rsDatCli
           BuscaCta = Replace(!Descripcio, "‚", "e", , , vbTextCompare)
           BuscaCta = Replace(!Descripcio, "¢", "o", , , vbTextCompare)
           BuscaCta = Replace(!Descripcio, "¡", "i", , , vbTextCompare)
           BuscaCta = Replace(!Descripcio, "§", "o.", , , vbTextCompare)
       End With
    End If
    rsDatCli.Close
    Set rsDatCli = Nothing
    
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Private Function BuscaIFI(ByVal psCodIFI As String) As String
    BuscaIFI = "NO INDICADO ..."
    Dim sCadSQl As String, sRutCta As String
    Dim rsDatCli As ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion

    Set rsDatCli = New ADODB.Recordset
    
    sCadSQl = "SELECT  Nombre  " _
            & "FROM   " & gsBaseRCC & "IFIs  " _
            & "WHERE Codigo =" & Int(psCodIFI)
    
    Set rsDatCli = oCon.CargaRecordSet(sCadSQl)
    If Not RSVacio(rsDatCli) Then
       With rsDatCli
           BuscaIFI = !Nombre
       End With
    End If
    rsDatCli.Close
    Set rsDatCli = Nothing
    
    oCon.CierraConexion
    Set oCon = Nothing
End Function


