VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCCECargaArch 
   BackColor       =   &H80000016&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carga de  Archivos Recibidos de la CCE"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7650
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCCECargaArch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLimpiar 
      Caption         =   "Limpiar"
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
      Left            =   6480
      TabIndex        =   10
      Top             =   6480
      Width           =   1095
   End
   Begin VB.TextBox txtResultado 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   3975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   2160
      Width           =   7095
   End
   Begin MSComDlg.CommonDialog dlgArchivo 
      Left            =   1920
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCargar 
      Caption         =   "Cargar"
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
      Left            =   3840
      TabIndex        =   5
      Top             =   6480
      Width           =   1230
   End
   Begin VB.Frame fraArchivoLOG 
      BackColor       =   &H80000016&
      Caption         =   "Cámara de Compensación Electrónica"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5265
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   7575
      Begin MSComctlLib.ProgressBar PbCCE 
         Height          =   255
         Left            =   1200
         TabIndex        =   8
         Top             =   720
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "..."
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
         Left            =   6720
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   3
         Top             =   360
         Width           =   465
      End
      Begin VB.TextBox txtCarpeta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   5280
      End
      Begin VB.Label lblPB 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6720
         TabIndex        =   9
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000016&
         Caption         =   "Estado     :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblRura 
         BackColor       =   &H80000016&
         Caption         =   "Ubicación :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   4
         Top             =   375
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
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
      Left            =   5160
      TabIndex        =   0
      Top             =   6480
      Width           =   1230
   End
   Begin VB.Frame FraHorario 
      BackColor       =   &H80000016&
      Caption         =   "Horario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   7575
      Begin VB.CommandButton Command1 
         Caption         =   "Refresh"
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
         Left            =   6600
         TabIndex        =   22
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblFinChe 
         BackColor       =   &H80000016&
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
         Left            =   6000
         TabIndex        =   21
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblIniChe 
         BackColor       =   &H80000016&
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
         Left            =   5160
         TabIndex        =   20
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblFinTra 
         BackColor       =   &H80000016&
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
         Left            =   6000
         TabIndex        =   19
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblIniTra 
         BackColor       =   &H80000016&
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
         Left            =   5160
         TabIndex        =   18
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblDesCheque 
         BackColor       =   &H80000016&
         Caption         =   "Cheque          :"
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
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblDesTranseferencia 
         BackColor       =   &H80000016&
         Caption         =   "Transferencia :"
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
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblCheDes 
         BackColor       =   &H0000C000&
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
         Height          =   255
         Left            =   2160
         TabIndex        =   15
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label lblChe 
         BackColor       =   &H0000C000&
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
         Height          =   255
         Left            =   1560
         TabIndex        =   14
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblTraDes 
         BackColor       =   &H0000C000&
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
         Height          =   255
         Left            =   2160
         TabIndex        =   13
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label lblTra 
         Alignment       =   2  'Center
         BackColor       =   &H0000C000&
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
         Height          =   255
         Left            =   1560
         TabIndex        =   12
         Top             =   240
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmCCECargaArch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************************************************
'** Nombre : frmCCECargaArch
'** Descripción : Para la Carga  de Archivos  , Proyecto: Implementacion del Servicio de Compensaciòn Electrónica Diferido de Instrumentos Compensables CCE
'** Creación : VAPA, 20160813
'*******************************************************************************************************************
Option Explicit
Dim oCCE As COMNCajaGeneral.NCOMCCE
Dim oCCER As COMNCajaGeneral.NCOMCCE
Dim oCamara As COMNCajaGeneral.NCOMCCE
Dim rsConstante As ADODB.Recordset
Dim gCCETransArch As String
Dim gCCEConTransArch As String
Dim gCCESalTransArch As String
Dim gCCEChequeArch As String
Dim gCCESalChequeArch As String
Dim gCCEConfLiqCheque As String
Dim Data() As String
Dim NBdata As Integer


Public Function LeerCarpeta(ByVal Carp As String) As Integer
Dim obj, f, F1, Chem As String
Dim CarpP As Folder
Dim i As Integer, Ext As String
Dim Ruta As String
Dim T As Double

    Set obj = CreateObject("Scripting.FileSystemObject")
    Set CarpP = obj.GetFolder(Carp)
    Chem = Carp: If Right(Ruta, 1) <> "\" Then Ruta = Ruta & "\"
    Set f = CarpP.Files
    GoSub RellenarData
    
Exit Function
RellenarData:
    For Each F1 In f
        Ext = LCase(Right(F1.Name, 3))
        If Ext = "txt" Then
            ReDim Preserve Data(1, NBdata)
            Data(0, NBdata) = F1.Name
            NBdata = NBdata + 1
        End If
    Next F1
    
Return
End Function
Private Sub cmdBuscar_Click()

Dim rsCarpeta As ADODB.Recordset
Dim i As Integer
Dim sTexto, lsRuta As String
Dim Linea, Dato, DataB() As String
Dim lnLoteMN, lnLoteME, lnRegMN, lnRegME As Long
Dim lbCuentaSoles As Boolean
Dim NumeroArchivo As Integer
Dim valor As Boolean
Dim myCadena As String
Dim FechaCreacion As Date
      txtResultado.Text = ""
     
    Set rsCarpeta = oCCE.CCE_ObtieneArchivosRecibidos
    txtCarpeta = rsCarpeta!nConsSisValor
    txtResultado.Text = txtResultado.Text & "Los Archivos encontrados son : " & vbCrLf
    txtResultado.Text = txtResultado.Text & vbCrLf
    
Erase Data
  'ReDim Preserve DataB(1, 0)
    If Len(txtCarpeta) > 1 Then
        LeerCarpeta (txtCarpeta)
        i = 0
            Do Until i = NBdata
                sTexto = Data(0, i)
                lsRuta = txtCarpeta.Text & "\" & sTexto
                i = i + 1
                If Not sTexto = "" Then
                    Select Case sTexto
                                        Case "CETRIAR.TXT"
                                               FechaCreacion = FileSystem.FileDateTime(lsRuta)
                                               txtResultado.Text = txtResultado.Text & "CETRIAR.TXT - Tranferencia Intermedia Abono  " & FechaCreacion & vbCrLf
                                        Case "CETRMAR.TXT"
                                               FechaCreacion = FileSystem.FileDateTime(lsRuta)
                                               txtResultado.Text = txtResultado.Text & "CETRMAR.TXT- Tranferencia Mañana Abono  " & FechaCreacion & vbCrLf
                                        Case "CETRTAR.TXT"
                                               FechaCreacion = FileSystem.FileDateTime(lsRuta)
                                               txtResultado.Text = txtResultado.Text & "CETRTAR.TXT- Tranferencia Tarde Abono  " & FechaCreacion & vbCrLf
                                        '-----------------------------------------------------------------------------------------------------------------------------
                                       Case "CETRIAR.txt"
                                               FechaCreacion = FileSystem.FileDateTime(lsRuta)
                                               txtResultado.Text = txtResultado.Text & "CETRIAR.txt - Tranferencia Intermedia Abono  " & FechaCreacion & vbCrLf
                                        Case "CETRMAR.txt"
                                               FechaCreacion = FileSystem.FileDateTime(lsRuta)
                                               txtResultado.Text = txtResultado.Text & "CETRMAR.txt- Tranferencia Mañana Abono  " & FechaCreacion & vbCrLf
                                        Case "CETRTAR.txt"
                                               FechaCreacion = FileSystem.FileDateTime(lsRuta)
                                               txtResultado.Text = txtResultado.Text & "CETRTAR.txt- Tranferencia Tarde Abono  " & FechaCreacion & vbCrLf
                                               
                                               
                                               
                                               
                                               
                                        Case "CETRIDR.TXT"
                                               FechaCreacion = FileSystem.FileDateTime(lsRuta)
                                               txtResultado.Text = txtResultado.Text & "CETRIDR.TXT- Tranferencia Intermedia Devolución  " & FechaCreacion & vbCrLf
                                        Case "CETRMDR.TXT"
                                               FechaCreacion = FileSystem.FileDateTime(lsRuta)
                                               txtResultado.Text = txtResultado.Text & "CETRMDR.TXT- Tranferencia Mañana Devolución  " & FechaCreacion & vbCrLf
                                        Case "CETRTDR.TXT"
                                               FechaCreacion = FileSystem.FileDateTime(lsRuta)
                                               txtResultado.Text = txtResultado.Text & "CETRTDR.TXT- Tranferencia Tarde Devolución  " & FechaCreacion & vbCrLf
                                               
                                        '-----------------------------------------------------------------------------------------------------------------------------
                                         Case "CETRIDR.txt"
                                               FechaCreacion = FileSystem.FileDateTime(lsRuta)
                                               txtResultado.Text = txtResultado.Text & "CETRIDR.txt- Tranferencia Intermedia Devolución  " & FechaCreacion & vbCrLf
                                        Case "CETRMDR.txt"
                                               FechaCreacion = FileSystem.FileDateTime(lsRuta)
                                               txtResultado.Text = txtResultado.Text & "CETRMDR.txt- Tranferencia Mañana Devolución  " & FechaCreacion & vbCrLf
                                        Case "CETRTDR.txt"
                                               FechaCreacion = FileSystem.FileDateTime(lsRuta)
                                               txtResultado.Text = txtResultado.Text & "CETRTDR.txt- Tranferencia Tarde Devolución  " & FechaCreacion & vbCrLf
                                               
                                               
                                               
                                               
                                               
'
                                        Case "CETRIPR.TXT"
                                               FechaCreacion = FileSystem.FileDateTime(lsRuta)
                                               txtResultado.Text = txtResultado.Text & "CETRIPR.TXT- Tranferencia Intermedia Presentados " & FechaCreacion & vbCrLf
                                        Case "CETRMPR.TXT"
                                               FechaCreacion = FileSystem.FileDateTime(lsRuta)
                                               txtResultado.Text = txtResultado.Text & "CETRMPR.TXT- Tranferencia Mañana Presentados " & FechaCreacion & vbCrLf
                                        Case "CETRTPR.TXT"
                                               FechaCreacion = FileSystem.FileDateTime(lsRuta)
                                               txtResultado.Text = txtResultado.Text & "CETRTPR.TXT- Tranferencia Tarde Presentados " & FechaCreacion & vbCrLf
                                               
                                        '-----------------------------------------------------------------------------------------------------------------------------
                                        Case "CETRIPR.txt"
                                               FechaCreacion = FileSystem.FileDateTime(lsRuta)
                                               txtResultado.Text = txtResultado.Text & "CETRIPR.txt- Tranferencia Intermedia Presentados " & FechaCreacion & vbCrLf
                                        Case "CETRMPR.txt"
                                               FechaCreacion = FileSystem.FileDateTime(lsRuta)
                                               txtResultado.Text = txtResultado.Text & "CETRMPR.txt- Tranferencia Mañana Presentados " & FechaCreacion & vbCrLf
                                        Case "CETRTPR.txt"
                                               FechaCreacion = FileSystem.FileDateTime(lsRuta)
                                               txtResultado.Text = txtResultado.Text & "CETRTPR.txt- Tranferencia Tarde Presentados " & FechaCreacion & vbCrLf
                                               
                                               
                                               
                                               
                                        Case "CETRISR.TXT"
                                               FechaCreacion = FileSystem.FileDateTime(lsRuta)
                                               txtResultado.Text = txtResultado.Text & "CETRISR.TXT- Tranferencia Intermedia Saldos " & FechaCreacion & vbCrLf
                                        Case "CETRMSR.TXT"
                                               FechaCreacion = FileSystem.FileDateTime(lsRuta)
                                               txtResultado.Text = txtResultado.Text & "CETRMSR.TXT- Tranferencia Mañana Saldos " & FechaCreacion & vbCrLf
                                        Case "CETRTSR.TXT"
                                               FechaCreacion = FileSystem.FileDateTime(lsRuta)
                                               txtResultado.Text = txtResultado.Text & "CETRTSR.TXT- Tranferencia Tarde Saldos " & FechaCreacion & vbCrLf
                                               
                                               
                                       '-----------------------------------------------------------------------------------------------------------------------------
                                       Case "CETRISR.txt"
                                               FechaCreacion = FileSystem.FileDateTime(lsRuta)
                                               txtResultado.Text = txtResultado.Text & "CETRISR.txt- Tranferencia Intermedia Saldos " & FechaCreacion & vbCrLf
                                        Case "CETRMSR.txt"
                                               FechaCreacion = FileSystem.FileDateTime(lsRuta)
                                               txtResultado.Text = txtResultado.Text & "CETRMSR.TXT- Tranferencia Mañana Saldos " & FechaCreacion & vbCrLf
                                        Case "CETRTSR.txt"
                                               FechaCreacion = FileSystem.FileDateTime(lsRuta)
                                               txtResultado.Text = txtResultado.Text & "CETRTSR.txt- Tranferencia Tarde Saldos " & FechaCreacion & vbCrLf
                                               
                                               
                                               
                                        Case "CECRECR.TXT"
                                               FechaCreacion = FileSystem.FileDateTime(lsRuta)
                                               txtResultado.Text = txtResultado.Text & "CECRECR.TXT- Cheques Rechazados  " & FechaCreacion & vbCrLf
                                       Case "CECSPVR.TXT"
                                               FechaCreacion = FileSystem.FileDateTime(lsRuta)
                                               txtResultado.Text = txtResultado.Text & "CECSPVR.TXT- Cheques Saldos Previos  " & FechaCreacion & vbCrLf
                                       Case "CECSALR.TXT"
                                               FechaCreacion = FileSystem.FileDateTime(lsRuta)
                                               txtResultado.Text = txtResultado.Text & "CECSALR.TXT- Cheques Saldos Finales  " & FechaCreacion & vbCrLf
                                      Case "CECSAUR.TXT"
                                               FechaCreacion = FileSystem.FileDateTime(lsRuta)
                                               txtResultado.Text = txtResultado.Text & "CECSAUR.TXT- Cheques Saldos Por RCG  " & FechaCreacion & vbCrLf
                                       Case "CECPRER.TXT"
                                               FechaCreacion = FileSystem.FileDateTime(lsRuta)
                                               txtResultado.Text = txtResultado.Text & "CECPRER.TXT- Cheques Presentados  " & FechaCreacion & vbCrLf
                                      '-----------------------------------------------------------------------------------------------------------------------------
                                      Case "CECRECR.txt"
                                               FechaCreacion = FileSystem.FileDateTime(lsRuta)
                                               txtResultado.Text = txtResultado.Text & "CECRECR.TXT- Cheques Rechazados  " & FechaCreacion & vbCrLf
                                       Case "CECSPVR.txt"
                                               FechaCreacion = FileSystem.FileDateTime(lsRuta)
                                               txtResultado.Text = txtResultado.Text & "CECSPVR.TXT- Cheques Saldos Previos  " & FechaCreacion & vbCrLf
                                       Case "CECSALR.txt"
                                               FechaCreacion = FileSystem.FileDateTime(lsRuta)
                                               txtResultado.Text = txtResultado.Text & "CECSALR.TXT- Cheques Saldos Finales  " & FechaCreacion & vbCrLf
                                      Case "CECSAUR.txt"
                                               FechaCreacion = FileSystem.FileDateTime(lsRuta)
                                               txtResultado.Text = txtResultado.Text & "CECSAUR.TXT- Cheques Saldos Por RCG  " & FechaCreacion & vbCrLf
                                       Case "CECPRER.txt"
                                               FechaCreacion = FileSystem.FileDateTime(lsRuta)
                                               txtResultado.Text = txtResultado.Text & "CECPRER.TXT- Cheques Presentados  " & FechaCreacion & vbCrLf

                    End Select
                    'txtResultado.Text = txtResultado.Text & sTexto & vbCrLf
                End If
            Loop
    Exit Sub
    Else
        txtResultado.Text = txtResultado.Text & _
        "No se encontraron archivos en la carpeta: " & vbCrLf
    Exit Sub
    End If
 End Sub

Private Sub LeerTramaRecibido()
Dim i As Integer
    If Len(txtCarpeta.Text) = 0 Then
        txtResultado.Text = txtResultado.Text & "No se seleccionó ninguna carpeta..." & vbCrLf
        Exit Sub
    Else
           Do Until i = NBdata
'''Lectura  de Archivos de Confirmacion de Abono Transferencias********
        If InStr(1, Data(0, i), "CETRIAR", vbTextCompare) <> 0 Then
                Leer_Confirmado_Recibido (txtCarpeta.Text & "\" & (Data(0, i)))
        End If
        If InStr(1, Data(0, i), "CETRMAR", vbTextCompare) <> 0 Then
                Leer_Confirmado_Recibido (txtCarpeta.Text & "\" & (Data(0, i)))
        End If
        If InStr(1, Data(0, i), "CETRTAR", vbTextCompare) <> 0 Then
                Leer_Confirmado_Recibido (txtCarpeta.Text & "\" & (Data(0, i)))
        End If
'''Lectura  de Archivos de Devolucion Transferencias******************
        If InStr(1, Data(0, i), "CETRIDR", vbTextCompare) <> 0 Then
                Leer_Devuelto_Recibido (txtCarpeta.Text & "\" & (Data(0, i)))
        End If
        If InStr(1, Data(0, i), "CETRMDR", vbTextCompare) <> 0 Then
                Leer_Devuelto_Recibido (txtCarpeta.Text & "\" & (Data(0, i)))
        End If
        If InStr(1, Data(0, i), "CETRTDR", vbTextCompare) <> 0 Then
                Leer_Devuelto_Recibido (txtCarpeta.Text & "\" & (Data(0, i)))
        End If
''Lectura de Archivos de Presentados Transferencias*****************
        If InStr(1, Data(0, i), "CETRIPR", vbTextCompare) <> 0 Then
                LeerPresentado_Recibido (txtCarpeta.Text & "\" & (Data(0, i)))
        End If
        If InStr(1, Data(0, i), "CETRMPR", vbTextCompare) <> 0 Then
                LeerPresentado_Recibido (txtCarpeta.Text & "\" & (Data(0, i)))
        End If
        If InStr(1, Data(0, i), "CETRTPR", vbTextCompare) <> 0 Then
                LeerPresentado_Recibido (txtCarpeta.Text & "\" & (Data(0, i)))
        End If
'''Lectura  de Archivos de Saldos Transferencias*****************
       If InStr(1, Data(0, i), "CETRISR", vbTextCompare) <> 0 Then
                Leer_Saldo_Recibido (txtCarpeta.Text & "\" & (Data(0, i)))
       End If
       If InStr(1, Data(0, i), "CETRMSR", vbTextCompare) <> 0 Then
                Leer_Saldo_Recibido (txtCarpeta.Text & "\" & (Data(0, i)))
       End If
       If InStr(1, Data(0, i), "CETRTSR", vbTextCompare) <> 0 Then
                Leer_Saldo_Recibido (txtCarpeta.Text & "\" & (Data(0, i)))
       End If
''''Lectura  de Presentados CHEQUES "NO RECIBIMOS PRESENTADOS " *********
       If InStr(1, Data(0, i), "CECPRER", vbTextCompare) <> 0 Then
                LeerCheque_Presentado_Recibido (txtCarpeta.Text & "\" & (Data(0, i)))
       End If
''''Lectura  de Rechazados CHEQUES **********************************
        If InStr(1, Data(0, i), "CECRECR", vbTextCompare) <> 0 Then
                LeerCheque_Rechazados_Recibido (txtCarpeta.Text & "\" & (Data(0, i)))
        End If
''''Lectura  de Conformidad de Liquidacion CHEQUES *******************
          If InStr(1, Data(0, i), "CECLIQR", vbTextCompare) <> 0 Then
                LeerCheque_Nulo_Liquidacion (txtCarpeta.Text & "\" & (Data(0, i)))
        End If
''''Lectura  de Saldos Previos CHEQUES *****************************
        If InStr(1, Data(0, i), "CECSPVR", vbTextCompare) <> 0 Then
                LeerCheque_SaldosPrevio_Recibido (txtCarpeta.Text & "\" & (Data(0, i)))
         End If
''''Lectura  de  Saldos por RCG CHEQUES *********************************
        If InStr(1, Data(0, i), "CECSAUR", vbTextCompare) <> 0 Then
                LeerCheque_SaldosRCG_Recibido (txtCarpeta.Text & "\" & (Data(0, i)))
          End If
''''Lectura  de Saldos Finales CHEQUES ********************************
        If InStr(1, Data(0, i), "CECSALR", vbTextCompare) <> 0 Then
                LeerCheque_SaldosFinales_Recibido (txtCarpeta.Text & "\" & (Data(0, i)))
         End If
                 i = i + 1
        Loop
    End If
End Sub
Private Sub Leer_Devuelto_Recibido(Optional ByRef psFile As String = "")
 Dim oFSO As Scripting.FileSystemObject
    Dim oStream As Scripting.TextStream
    Dim mvarRuta, sNombreArchivo, sRecibido, sId, sFechaArchivo, sCodAplicacion, sNumArchivo, sUsuario, sMoneda, Linea, Data, lsCabDataMN, lsCabDataME As String
    Dim nId, NumeroArchivo As Integer
    Dim lnLoteMN, lnLoteME, lnRegMN, lnRegME As Long
    Dim lbCuentaSoles As Boolean
    Dim lvConfirmado() As String
    Dim lvDevuelto() As String
    Dim lsCabLote, lsCabArchivo, lsRegIndividual, lsRegAdicional, sCCI, sDoi, sVal, sEntAcre, espacios, espacios2, sSecU, sContReg As String
    Dim lnIDMN, lnIDME As Long
    Dim lnIDRef As Long
    Dim sFechaMB As String
    Dim lsMovNro, lsSecUnivoca, lsTipTra, lsCabNullMN, lsCabNullME As String
    Dim lbArch As Boolean
    
'cambioultimo
On Error GoTo ErrorDev
oCCE.BeginTrans
lbArch = True


     Set rsConstante = oCCE.CCE_ConstArchivos
     gCCETransArch = rsConstante!gCCETransArch
     
    If Len(psFile) <> 0 Then mvarRuta = psFile
    If Len(mvarRuta) = 0 Then Exit Sub: oCCE.CommitTrans
    
    Set oFSO = New Scripting.FileSystemObject
    Set oStream = oFSO.OpenTextFile(mvarRuta, ForReading)
    sNombreArchivo = oFSO.GetFileName(mvarRuta)
    NumeroArchivo = FreeFile
    'Controles **********************************
    Open mvarRuta For Input As #NumeroArchivo
    While Not EOF(NumeroArchivo)
            Line Input #NumeroArchivo, Linea
            Data = Linea
            Select Case UCase(Left(Data, 1))
                    Case 1
                            If Mid(Data, 3, 1) = 1 Then
                                lbCuentaSoles = True
                                lnLoteMN = 0
                                lnRegMN = 0
                            Else
                                lbCuentaSoles = False
                                lnLoteME = 0
                                lnRegME = 0
                            End If
                    Case 5
                            If lbCuentaSoles Then
                                lnLoteMN = lnLoteMN + 1
                            Else
                                lnLoteME = lnLoteME + 1
                            End If
            End Select
            If lbCuentaSoles Then
                lnRegMN = lnRegMN + 1
            Else
                lnRegME = lnRegME + 1
            End If
    Wend
    Close #NumeroArchivo
    '***********************************************
    'Detalle***********************************
    lnIDMN = 0
    lnIDME = 0
    Open mvarRuta For Input As #NumeroArchivo
     While Not EOF(NumeroArchivo)
            Line Input #NumeroArchivo, Linea
            Data = Linea
            Select Case UCase(Left(Data, 1))
                    Case 1
                            sMoneda = Mid(Data, 3, 1)
                            sFechaArchivo = Mid(Data, 23, 8)
                            sCodAplicacion = Mid(Data, 4, 3)
                            sNumArchivo = CInt(Mid(Data, 31, 2))
                            If sMoneda = "1" Then
                                lsCabDataMN = oCCE.CCE_ValidaCabecera(Data)
                                lsCabNullMN = oCCE.CCE_ValidaCabeceraNulas(gCCETransArch, sFechaArchivo, sMoneda, sCodAplicacion, 12)
                                If lsCabDataMN = "no" And lsCabNullMN = "no" Then
                                    lnIDMN = oCCE.CCE_InsIntercambioTransferenciaRec(gCCETransArch, sFechaArchivo, sMoneda, sCodAplicacion, sNumArchivo, lnLoteMN, lnRegMN, 1, gsCodUser, 12)
                                Else
                                    sFechaMB = Right(sFechaArchivo, 2) & "/" & Mid(sFechaArchivo, 5, 2) & "/" & Left(sFechaArchivo, 4)
                                    MsgBox "Ya se realizo la Carga de Devolución de Transferencias del Archivo  " & sNombreArchivo & " ""   del :  " & sFechaMB & " ", vbInformation, "Aviso"
                                Close #NumeroArchivo
                                oCCE.CommitTrans
                                Exit Sub
                                End If
                            Else
                                lsCabDataME = oCCE.CCE_ValidaCabecera(Data)
                                lsCabNullME = oCCE.CCE_ValidaCabeceraNulas(gCCETransArch, sFechaArchivo, sMoneda, sCodAplicacion, 12)
                                If lsCabDataME = "no" And lsCabNullME = "no" Then
                                    lnIDME = oCCE.CCE_InsIntercambioTransferenciaRec(gCCETransArch, sFechaArchivo, sMoneda, sCodAplicacion, sNumArchivo, lnLoteME, lnRegME, 1, gsCodUser, 12)
                                Else
                                 sFechaMB = Right(sFechaArchivo, 2) & "/" & Mid(sFechaArchivo, 5, 2) & "/" & Left(sFechaArchivo, 4)
                                 MsgBox "Ya se realizo la Carga de Devolución de Transferencias del Archivo  " & sNombreArchivo & " ""   del :  " & sFechaMB & " ", vbInformation, "Aviso"
                                Close #NumeroArchivo
                                oCCE.CommitTrans
                                Exit Sub
                                End If
                            End If
            End Select
            If lnIDMN > 0 And sMoneda = "1" And lnRegMN <> 2 Then
                If Left(Data, 1) = "5" Then
                    lsTipTra = Mid(Data, 67, 3)
                End If
                If Left(Data, 1) = "6" Then
                   lsSecUnivoca = Mid(Data, 178, 7)
                   oCCE.CCE_RechazaTRA_Enviada lsSecUnivoca, gdFecSis
                End If
                oCCE.CCE_InsertaIntercambioTransferenciaDet lnIDMN, IIf(Left(Data, 1) = "6" Or Left(Data, 1) = "7", lsSecUnivoca, 0), UCase(Left(Data, 1)), Data, IIf(Left(Data, 1) = "5" Or Left(Data, 1) = "6" Or Left(Data, 1) = "7" Or Left(Data, 1) = "8", lsTipTra, 0), 12
            End If
            If lnIDME > 0 And sMoneda = "2" And lnRegME <> 2 Then
                If Left(Data, 1) = "5" Then
                    lsTipTra = Mid(Data, 67, 3)
                End If
                If Left(Data, 1) = "6" Then
                   lsSecUnivoca = Mid(Data, 178, 7)
                   oCCE.CCE_RechazaTRA_Enviada lsSecUnivoca, gdFecSis
                   
                End If
                oCCE.CCE_InsertaIntercambioTransferenciaDet lnIDME, IIf(Left(Data, 1) = "6" Or Left(Data, 1) = "7", lsSecUnivoca, 0), UCase(Left(Data, 1)), Data, IIf(Left(Data, 1) = "5" Or Left(Data, 1) = "6" Or Left(Data, 1) = "7" Or Left(Data, 1) = "8", lsTipTra, 0), 12
            End If
    Wend
    Close #NumeroArchivo
     txtResultado.Text = txtResultado.Text & _
                        "Archivo Cargado: " & sNombreArchivo & vbCrLf & _
                        "Registros Cargados: " & (lnRegMN + lnRegME) & vbCrLf
 oCCE.CommitTrans
    lbArch = False
    Exit Sub

ErrorDev:
    If lbArch Then
        oCCE.RollbackTrans
        Set oCCE = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Private Sub Leer_Confirmado_Recibido(Optional ByRef psFile As String = "")
    Dim oFSO As Scripting.FileSystemObject
    Dim oStream As Scripting.TextStream
    Dim mvarRuta, sNombreArchivo, sRecibido, sId, sFechaArchivo, sCodAplicacion, sNumArchivo, sUsuario, sMoneda, Linea, Data, lsCabDataMN, lsCabDataME As String
    Dim nId, NumeroArchivo As Integer
    Dim lnLoteMN, lnLoteME, lnRegMN, lnRegME As Long
    Dim lbCuentaSoles As Boolean
    Dim lvConfirmado() As String
    Dim lvDevuelto() As String
    Dim lsCabLote, lsCabArchivo, lsRegIndividual, lsRegAdicional, sCCI, sDoi, sVal, sEntAcre, espacios, espacios2, sSecU, sContReg As String
    Dim lnIDMN, lnIDME, lsSecUnivoca As Long
    Dim lnIDRef As Long
    Dim lsMovNro, lsSecUnivocaS, lsSecUnivocaD, lsTipTra, lsCabNullMN, lsCabNullME   As String
    Dim lbArch As Boolean
    Dim lsCtaCodCci As String
    Dim lsImporte As Double
    Dim lsImporteComision As Double
    Dim lsMismoTitular As String
    Dim nSubTpo As Integer
    Dim sFechaMB As String
    Dim lsConfirmacion As Long 'vapa20170113

  On Error GoTo ErrorCon
  oCCE.BeginTrans
  lbArch = True
    
    Set rsConstante = oCCE.CCE_ConstArchivos
     gCCEConTransArch = rsConstante!gCCEConTransArch
            
    If Len(psFile) <> 0 Then mvarRuta = psFile
    If Len(mvarRuta) = 0 Then Exit Sub: oCCE.CommitTrans
    
    Set oFSO = New Scripting.FileSystemObject
    Set oStream = oFSO.OpenTextFile(mvarRuta, ForReading)
    sNombreArchivo = oFSO.GetFileName(mvarRuta)
    NumeroArchivo = FreeFile
    'Controles **********************************
    Open mvarRuta For Input As #NumeroArchivo
    While Not EOF(NumeroArchivo)
            Line Input #NumeroArchivo, Linea
            Data = Linea
            Select Case UCase(Left(Data, 1))
                    Case 1
                            If Mid(Data, 3, 1) = 1 Then
                                lbCuentaSoles = True
                                lnLoteMN = 0
                                lnRegMN = 0
                            Else
                                lbCuentaSoles = False
                                lnLoteME = 0
                                lnRegME = 0
                            End If
                    Case 5
                            If lbCuentaSoles Then
                                lnLoteMN = lnLoteMN + 1
                            Else
                                lnLoteME = lnLoteME + 1
                            End If
            End Select
            If lbCuentaSoles Then
                lnRegMN = lnRegMN + 1
            Else
                lnRegME = lnRegME + 1
            End If
    Wend
    Close #NumeroArchivo
    '***********************************************
    'Detalle ***********************************
    lnIDMN = 0
    lnIDME = 0
    Open mvarRuta For Input As #NumeroArchivo
     While Not EOF(NumeroArchivo)
            Line Input #NumeroArchivo, Linea
            Data = Linea
            Select Case UCase(Left(Data, 1))
                    Case 1
                            sMoneda = Mid(Data, 3, 1)
                            sFechaArchivo = Mid(Data, 23, 8)
                            sCodAplicacion = Mid(Data, 4, 3)
                            sNumArchivo = CInt(Mid(Data, 31, 2))
                            
                            If sMoneda = "1" Then
                                lsCabDataMN = oCCE.CCE_ValidaCabecera(Data)
                                lsCabNullMN = oCCE.CCE_ValidaCabeceraNulas(gCCEConTransArch, sFechaArchivo, sMoneda, sCodAplicacion, 13)
                                If lsCabDataMN = "no" And lsCabNullMN = "no" Then
                                lnIDMN = oCCE.CCE_InsIntercambioTransferenciaRec(gCCEConTransArch, sFechaArchivo, sMoneda, sCodAplicacion, sNumArchivo, lnLoteMN, lnRegMN, "1", gsCodUser, 13)
                                Else
                                 sFechaMB = Right(sFechaArchivo, 2) & "/" & Mid(sFechaArchivo, 5, 2) & "/" & Left(sFechaArchivo, 4)
                                 MsgBox "Ya se realizo la Carga de Confirmados de Transferencias del Archivo  " & sNombreArchivo & " ""   del :  " & sFechaMB & " ", vbInformation, "Aviso"
                                Close #NumeroArchivo
                                oCCE.CommitTrans
                                Exit Sub
                                End If
                            Else
                                 lsCabDataME = oCCE.CCE_ValidaCabecera(Data)
                                 lsCabNullME = oCCE.CCE_ValidaCabeceraNulas(gCCEConTransArch, sFechaArchivo, sMoneda, sCodAplicacion, 13)
                                 If lsCabDataME = "no" And lsCabNullME = "no" Then
                                 lnIDME = oCCE.CCE_InsIntercambioTransferenciaRec(gCCEConTransArch, sFechaArchivo, sMoneda, sCodAplicacion, sNumArchivo, lnLoteME, lnRegME, "1", gsCodUser, 13)
                                 Else
                                  sFechaMB = Right(sFechaArchivo, 2) & "/" & Mid(sFechaArchivo, 5, 2) & "/" & Left(sFechaArchivo, 4)
                                 MsgBox "Ya se realizo la Carga de Confirmados de Transferencias del Archivo  " & sNombreArchivo & " ""   del :  " & sFechaMB & " ", vbInformation, "Aviso"
                                 Close #NumeroArchivo
                                 oCCE.CommitTrans
                                 Exit Sub
                                 End If
                            End If
            End Select
            If lnIDMN > 0 And sMoneda = "1" And lnRegMN <> 2 Then
                 If Left(Data, 1) = "5" Then
                    lsTipTra = Mid(Data, 67, 3)
                End If
                If Left(Data, 1) = "6" Then
                    lsSecUnivoca = CDbl(Mid(Data, 178, 7))
                    'vapa20170113
                  'VAPA20170711 COMENTADO
                   'lsConfirmacion = oCCE.CCE_ConfirmaTRA_Enviada(lsSecUnivoca, gdFecSis)  'confirmacion TRAMA comentado por vapa   wivapas
                   
                  ' If lsConfirmacion = 0 Then
                   '                 MsgBox "no se pudo hacer la confirmacion de abona de la transaccion  " & lsSecUnivoca & " ""   del :  " & sFechaMB & " ", vbInformation, "Aviso"
                    '                 oCCE.RollbackTrans
                     '               Exit Sub
                   'End If
                End If
                oCCE.CCE_InsertaIntercambioTransferenciaDet lnIDMN, IIf(Left(Data, 1) = "6" Or Left(Data, 1) = "7", lsSecUnivoca, 0), UCase(Left(Data, 1)), Data, IIf(Left(Data, 1) = "5" Or Left(Data, 1) = "6" Or Left(Data, 1) = "7" Or Left(Data, 1) = "8", lsTipTra, 0), 13
            End If
            If lnIDME > 0 And sMoneda = "2" And lnRegME <> 2 Then
                  If Left(Data, 1) = "5" Then
                    lsTipTra = Mid(Data, 67, 3)
                End If
                If Left(Data, 1) = "6" Then
                   lsSecUnivoca = CDbl(Mid(Data, 178, 7))
                  'VAPA20170811
'                 lsConfirmacion = oCCE.CCE_ConfirmaTRA_Enviada(lsSecUnivoca, gdFecSis)
'
'                  If lsConfirmacion = 0 Then
'                                MsgBox "no se pudo hacer la confirmacion de abona de la transaccion  " & lsSecUnivoca & " ""   del :  " & sFechaMB & " ", vbInformation, "Aviso"
'                                oCCE.RollbackTrans
'                   Exit Sub
'                   End If
                End If
                oCCE.CCE_InsertaIntercambioTransferenciaDet lnIDME, IIf(Left(Data, 1) = "6" Or Left(Data, 1) = "7", lsSecUnivoca, 0), UCase(Left(Data, 1)), Data, IIf(Left(Data, 1) = "5" Or Left(Data, 1) = "6" Or Left(Data, 1) = "7" Or Left(Data, 1) = "8", lsTipTra, 0), 13
            End If
    Wend
    Close #NumeroArchivo
    txtResultado.Text = txtResultado.Text & _
                        "Archivo Cargado: " & sNombreArchivo & vbCrLf & _
                        "Registros Cargados: " & (lnRegMN + lnRegME) & vbCrLf
oCCE.CommitTrans
    lbArch = False
    Exit Sub

ErrorCon:
    If lbArch Then
        oCCE.RollbackTrans
        Set oCCE = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub LeerPresentado_Recibido(Optional ByRef psFile As String = "")
    Dim oFSO As Scripting.FileSystemObject
    Dim oStream As Scripting.TextStream
    Dim mvarRuta, sNombreArchivo, sRecibido, sId, sFechaArchivo, sCodAplicacion, sNumArchivo, sUsuario, sMoneda, Linea, Data, lsCabDataMN, lsCabDataME As String
    Dim nId, NumeroArchivo As Integer
    Dim lnLoteMN, lnLoteME, lnRegMN, lnRegME As Long
    Dim lbCuentaSoles As Boolean
    Dim lvConfirmado() As String
    Dim lvDevuelto() As String
    Dim lsCabLote, lsCabArchivo, lsRegIndividual, lsRegAdicional, sCCI, sDoi, sVal, sEntAcre, espacios, espacios2, sSecU, sContReg As String
    Dim lnIDMN, lnIDME, lsCodSecBCR   As Long
    Dim lnIDRef As Long
    Dim lsMovNro, lsSecUnivoca, lsTipTra, lsCodBCR, lsCabNullMN, lsCabNullME As String
    Dim lbArch As Boolean
    Dim sFechaMB As String
    Dim dFormatFecha As String

On Error GoTo ErrorPreTransf
oCCE.BeginTrans
lbArch = True

      Set rsConstante = oCCE.CCE_ConstArchivos
      gCCETransArch = rsConstante!gCCETransArch
        
    If Len(psFile) <> 0 Then mvarRuta = psFile
    If Len(mvarRuta) = 0 Then Exit Sub: oCCE.CommitTrans
    
    Set oFSO = New Scripting.FileSystemObject
    Set oStream = oFSO.OpenTextFile(mvarRuta, ForReading)
    sNombreArchivo = oFSO.GetFileName(mvarRuta)
    
    
    NumeroArchivo = FreeFile
    'Controles **********************************
    Open mvarRuta For Input As #NumeroArchivo
    While Not EOF(NumeroArchivo)
            Line Input #NumeroArchivo, Linea
            Data = Linea
            Select Case UCase(Left(Data, 1))
                    Case 1
                            If Mid(Data, 3, 1) = 1 Then
                                lbCuentaSoles = True
                                lnLoteMN = 0
                                lnRegMN = 0
                            Else
                                lbCuentaSoles = False
                                lnLoteME = 0
                                lnRegME = 0
                            End If
                    Case 5
                            If lbCuentaSoles Then
                                lnLoteMN = lnLoteMN + 1
                            Else
                                lnLoteME = lnLoteME + 1
                            End If
            End Select
            If lbCuentaSoles Then
                lnRegMN = lnRegMN + 1
            Else
                lnRegME = lnRegME + 1
            End If
    Wend
    Close #NumeroArchivo
    '***********************************************
    'Detalle ***********************************
    lnIDMN = 0
    lnIDME = 0
    Open mvarRuta For Input As #NumeroArchivo
     While Not EOF(NumeroArchivo)
            Line Input #NumeroArchivo, Linea
            Data = Linea
            Select Case UCase(Left(Data, 1))
                    Case 1
                            sMoneda = Mid(Data, 3, 1)
                            sFechaArchivo = Mid(Data, 23, 8)
                            sCodAplicacion = Mid(Data, 4, 3)
                            sNumArchivo = CInt(Mid(Data, 31, 2))
                            
                            
                            If sMoneda = "1" Then
                                lsCabDataMN = oCCE.CCE_ValidaCabecera(Data)
                                lsCabNullMN = oCCE.CCE_ValidaCabeceraNulas(gCCETransArch, sFechaArchivo, sMoneda, sCodAplicacion, 11)
                                If lsCabDataMN = "no" And lsCabNullMN = "no" Then
                                lnIDMN = oCCE.CCE_InsIntercambioTransferenciaRec(gCCETransArch, sFechaArchivo, sMoneda, sCodAplicacion, sNumArchivo, lnLoteMN, lnRegMN, 1, gsCodUser, 11)
                                Else
                                    sFechaMB = Right(sFechaArchivo, 2) & "/" & Mid(sFechaArchivo, 5, 2) & "/" & Left(sFechaArchivo, 4)
                                    MsgBox "Ya se realizo la Carga de Presentados de Transferencias del Archivo  " & sNombreArchivo & " ""   del :  " & sFechaMB & " ", vbInformation, "Aviso"
                                Close #NumeroArchivo
                                oCCE.CommitTrans
                                Exit Sub
                                End If
                            Else
                                lsCabDataME = oCCE.CCE_ValidaCabecera(Data)
                                lsCabNullME = oCCE.CCE_ValidaCabeceraNulas(gCCETransArch, sFechaArchivo, sMoneda, sCodAplicacion, 11)
                                If lsCabDataME = "no" And lsCabNullME = "no" Then
                                    lnIDME = oCCE.CCE_InsIntercambioTransferenciaRec(gCCETransArch, sFechaArchivo, sMoneda, sCodAplicacion, sNumArchivo, lnLoteME, lnRegME, 1, gsCodUser, 11)
                                Else
                                        sFechaMB = Right(sFechaArchivo, 2) & "/" & Mid(sFechaArchivo, 5, 2) & "/" & Left(sFechaArchivo, 4)
                                        MsgBox "Ya se realizo la Carga de Presentados de Transferencias del Archivo  " & sNombreArchivo & " ""   del :  " & sFechaMB & " ", vbInformation, "Aviso"
                                Close #NumeroArchivo
                                oCCE.CommitTrans
                                Exit Sub
                                End If
                            End If
                            
            End Select
            
            If lnIDMN > 0 And sMoneda = "1" And lnRegMN <> 2 Then
                If Left(Data, 1) = "5" Then
                    lsTipTra = Mid(Data, 67, 3)
                End If
                If Left(Data, 1) = "6" Then
'                    lsMovNro = clsCap.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
'                    lnIDRef = clsCap.CCE_RegistraRefGlobal(lsMovNro)
                     lsSecUnivoca = Mid(Data, 178, 7)
                     'lsCodBCR = Mid(Data, 186, 4)
                     lsCodSecBCR = CLng(lsSecUnivoca)
                End If
                oCCE.CCE_InsertaIntercambioTransferenciaDet lnIDMN, IIf(Left(Data, 1) = "6" Or Left(Data, 1) = "7", lsCodSecBCR, 0), IIf(UCase(Left(Data, 1)) = "7", "71", UCase(Left(Data, 1))), Data, IIf(Left(Data, 1) = "5" Or Left(Data, 1) = "6" Or Left(Data, 1) = "7" Or Left(Data, 1) = "8", lsTipTra, 0), 11
                
            End If
            
            If lnIDME > 0 And sMoneda = "2" And lnRegME <> 2 Then
                If Left(Data, 1) = "5" Then
                    lsTipTra = Mid(Data, 67, 3)
                End If
                If Left(Data, 1) = "6" Then
'                    lsMovNro = clsCap.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
'                    lnIDRef = clsCap.CCE_RegistraRefGlobal(lsMovNro)
                     lsSecUnivoca = Mid(Data, 178, 7)
                     'lsCodBCR = Mid(Data, 186, 4)
                     lsCodSecBCR = CLng(lsSecUnivoca)
                End If
                oCCE.CCE_InsertaIntercambioTransferenciaDet lnIDME, IIf(Left(Data, 1) = "6" Or Left(Data, 1) = "7", lsCodSecBCR, 0), IIf(UCase(Left(Data, 1)) = "7", "71", UCase(Left(Data, 1))), Data, IIf(Left(Data, 1) = "5" Or Left(Data, 1) = "6" Or Left(Data, 1) = "7" Or Left(Data, 1) = "8", lsTipTra, 0), 11
                
            End If
    Wend
    Close #NumeroArchivo
    'Evaluacion de Presentados
    If (lnRegMN <> 2 And lnRegME <> 2) Or (lnRegMN <> 2 And lnRegME = 2) Or (lnRegMN = 2 And lnRegME <> 2) Then
    
    dFormatFecha = Format(gdFecSis, "YYYYmmdd")
                    
                        oCCE.CCE_ValidaPresentadosRec sFechaArchivo, sCodAplicacion 'vapa20170707
                  
    End If
'    If lnIDMN > 0 Then oCCE.CCE_RegistraTRA_Validado lnIDMN
'    If lnIDME > 0 Then oCCE.CCE_RegistraTRA_Validado lnIDME
  ' oCCE.CCE_VALIDACION
  
 ' oCCE.CCE_ValidaPresentadosRec ' ya no invocarlo
  
      txtResultado.Text = txtResultado.Text & _
                        "Archivo Cargado: " & sNombreArchivo & vbCrLf & _
                        "Registros Cargados: " & (lnRegMN + lnRegME) & vbCrLf
 oCCE.CommitTrans
    lbArch = False
    Exit Sub

ErrorPreTransf:
    If lbArch Then
        oCCE.RollbackTrans
        Set oCCE = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Private Sub Leer_Saldo_Recibido(Optional ByRef psFile As String = "")
     Dim oFSO As Scripting.FileSystemObject
    Dim oStream As Scripting.TextStream
    Dim mvarRuta, sNombreArchivo, sRecibido, sId, sFechaArchivo, sCodAplicacion, sNumArchivo, sUsuario, sMoneda, Linea, Data, lsCabDataMN, lsCabDataME As String
    Dim nId, NumeroArchivo As Integer
    Dim lnLote, lnReg As Long
    Dim lbCuentaSoles As Boolean
    Dim lvConfirmado() As String
    Dim lvDevuelto() As String
    Dim lsCabLote, lsCabArchivo, lsRegIndividual, lsRegAdicional, sCCI, sDoi, sVal, sEntAcre, espacios, espacios2, sSecU, sContReg As String
    Dim lnIDMN, lnIDME As Long
    Dim lnIDRef As Long
    Dim lsMovNro, lsSecUnivoca, lsCabNullMN As String
    Dim lbArch As Boolean
    Dim sFechaMB As String
    
On Error GoTo ErrorSal
oCCE.BeginTrans
lbArch = True

      Set rsConstante = oCCE.CCE_ConstArchivos
      gCCESalTransArch = rsConstante!gCCESalTransArch
        
    If Len(psFile) <> 0 Then mvarRuta = psFile
    If Len(mvarRuta) = 0 Then Exit Sub: oCCE.CommitTrans
    
    Set oFSO = New Scripting.FileSystemObject
    Set oStream = oFSO.OpenTextFile(mvarRuta, ForReading)
    sNombreArchivo = oFSO.GetFileName(mvarRuta)
    NumeroArchivo = FreeFile
    'Controles **********************************
    Open mvarRuta For Input As #NumeroArchivo
    While Not EOF(NumeroArchivo)
            Line Input #NumeroArchivo, Linea
            Data = Linea
            Select Case UCase(Left(Data, 2))
                    Case "CA"
                                lnLote = 0
                                lnReg = 0
                    Case "CL"
                            sMoneda = Mid(Data, 7, 1)
                            lnLote = lnLote + 1
            End Select
               lnReg = lnReg + 1
            
    Wend
    Close #NumeroArchivo
    
    
    '***********************************************
    'Detalle********************
    lnIDMN = 0
    lnIDME = 0
    Open mvarRuta For Input As #NumeroArchivo
     While Not EOF(NumeroArchivo)
            Line Input #NumeroArchivo, Linea
            Data = Linea
            Select Case UCase(Left(Data, 2))
                    Case "CA"
                            sFechaArchivo = Mid(Data, 43, 8)
                            sCodAplicacion = Mid(Data, 19, 3)
                            sNumArchivo = CInt(Mid(Data, 41, 1))
                            lsCabDataMN = oCCE.CCE_ValidaCabecera(Data)
                            lsCabNullMN = oCCE.CCE_ValidaCabeceraNulas(gCCESalTransArch, sFechaArchivo, 0, sCodAplicacion, 17)
                                If lsCabDataMN = "no" And lsCabNullMN = "no" Then
                                lnIDMN = oCCE.CCE_InsIntercambioTransferenciaRec(gCCESalTransArch, sFechaArchivo, 0, sCodAplicacion, sNumArchivo, lnLote, lnReg, 1, gsCodUser, 17)
                                Else
                                    sFechaMB = Right(sFechaArchivo, 2) & "/" & Mid(sFechaArchivo, 5, 2) & "/" & Left(sFechaArchivo, 4)
                                    MsgBox "Ya se realizo la Carga de Saldos de Transferencias del Archivo  " & sNombreArchivo & " ""   del :  " & sFechaMB & " ", vbInformation, "Aviso"
                                Close #NumeroArchivo
                                oCCE.CommitTrans
                                Exit Sub
                                End If
         End Select
            If lnIDMN > 0 And lnReg <> 2 Then
                oCCE.CCE_InsertaIntercambioTransferenciaDet lnIDMN, 0, UCase(Left(Data, 2)), Data, "SAL", 17
            End If
    Wend
    Close #NumeroArchivo

       txtResultado.Text = txtResultado.Text & _
                        "Archivo Cargado: " & sNombreArchivo & vbCrLf & _
                        "Registros Cargados: " & lnReg & vbCrLf
 oCCE.CommitTrans
    lbArch = False
    Exit Sub

ErrorSal:
    If lbArch Then
        oCCE.RollbackTrans
        Set oCCE = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub LeerCheque_Presentado_Recibido(Optional ByRef psFile As String = "")
    Dim oFSO As Scripting.FileSystemObject
    Dim oStream As Scripting.TextStream
    Dim mvarRuta, sNombreArchivo, sRecibido, sId, sFechaArchivo, sCodAplicacion, sNumArchivo, sUsuario, sMoneda, Linea, Data, lsCabDataMN, lsCabDataME As String
    Dim nId, NumeroArchivo As Integer
    Dim lnLoteMN, lnLoteME, lnRegMN, lnRegME As Long
    Dim lbCuentaSoles As Boolean
    Dim lvConfirmado() As String
    Dim lvDevuelto() As String
    Dim lsCabLote, lsCabArchivo, lsRegIndividual, lsRegAdicional, sCCI, sDoi, sVal, sEntAcre, espacios, espacios2, sSecU, sContReg As String
    Dim lnIDMN, lnIDME As Long
    Dim lnIDRef As Long
    Dim lsMovNro, lsSecUnivoca, lsTipTra As String
    Dim lbArch As Boolean
    Dim sFechaMB As String

On Error GoTo ErrorCheP
oCCE.BeginTrans
lbArch = True
    
    Set rsConstante = oCCE.CCE_ConstArchivos
    gCCEChequeArch = rsConstante!gCCEChequeArch
        
        
    If Len(psFile) <> 0 Then mvarRuta = psFile
    If Len(mvarRuta) = 0 Then Exit Sub: oCCE.CommitTrans
    
    Set oFSO = New Scripting.FileSystemObject
    Set oStream = oFSO.OpenTextFile(mvarRuta, ForReading)
    sNombreArchivo = oFSO.GetFileName(mvarRuta)
    NumeroArchivo = FreeFile
    'Controles **********************************
    Open mvarRuta For Input As #NumeroArchivo
    While Not EOF(NumeroArchivo)
            Line Input #NumeroArchivo, Linea
            Data = Linea
            Select Case UCase(Left(Data, 1))
                    Case 1
                            If Mid(Data, 3, 1) = 1 Then
                                lbCuentaSoles = True
                                lnLoteMN = 0
                                lnRegMN = 0
                            Else
                                lbCuentaSoles = False
                                lnLoteME = 0
                                lnRegME = 0
                            End If
                    Case 5
                            If lbCuentaSoles Then
                                lnLoteMN = lnLoteMN + 1
                            Else
                                lnLoteME = lnLoteME + 1
                            End If
            End Select
            If lbCuentaSoles Then
                lnRegMN = lnRegMN + 1
            Else
                lnRegME = lnRegME + 1
            End If
    Wend
    Close #NumeroArchivo
    '*******************************************************************
    'Detalle*************
    lnIDMN = 0
    lnIDME = 0
    Open mvarRuta For Input As #NumeroArchivo
     While Not EOF(NumeroArchivo)
            Line Input #NumeroArchivo, Linea
            Data = Linea
            Select Case UCase(Left(Data, 1))
                    Case 1
                            sMoneda = Mid(Data, 3, 1)
                            sFechaArchivo = Mid(Data, 23, 8)
                            sCodAplicacion = Mid(Data, 4, 3)
                            sNumArchivo = CInt(Mid(Data, 31, 2))
                            If sMoneda = "1" Then
                                lsCabDataMN = oCCE.CCE_ValidaCabecera(Data)
                                If lsCabDataMN = "no" Then
                                lnIDMN = oCCE.CCE_InsIntercambioTransferenciaRec(gCCEChequeArch, sFechaArchivo, sMoneda, sCodAplicacion, sNumArchivo, lnLoteMN, lnRegMN, 1, gsCodUser, 11)
                                Else
                                sFechaMB = Right(sFechaArchivo, 2) & "/" & Mid(sFechaArchivo, 5, 2) & "/" & Left(sFechaArchivo, 4)
                                MsgBox "Ya se realizo la Carga de Presentados de Cheques del Archivo  " & sNombreArchivo & " ""   del :  " & sFechaMB & " ", vbInformation, "Aviso"
                                Close #NumeroArchivo
                                oCCE.CommitTrans
                                Exit Sub
                                End If
                            Else
                                lsCabDataME = oCCE.CCE_ValidaCabecera(Data)
                                If lsCabDataME = "no" Then
                                lnIDME = oCCE.CCE_InsIntercambioTransferenciaRec(gCCEChequeArch, sFechaArchivo, sMoneda, sCodAplicacion, sNumArchivo, lnLoteME, lnRegME, 1, gsCodUser, 11)
                                Else
                                sFechaMB = Right(sFechaArchivo, 2) & "/" & Mid(sFechaArchivo, 5, 2) & "/" & Left(sFechaArchivo, 4)
                                MsgBox "Ya se realizo la Carga de Presentados de Cheques del Archivo  " & sNombreArchivo & " ""   del :  " & sFechaMB & " ", vbInformation, "Aviso"
                                Close #NumeroArchivo
                                oCCE.CommitTrans
                                Exit Sub
                                End If
                            End If
            End Select
            If lnIDMN > 0 And sMoneda = "1" And lnRegMN <> 2 Then
                If Left(Data, 1) = "5" Then
                   ' lsTipTra = Mid(Data, 67, 3)
                End If
                If Left(Data, 1) = "6" Then
                   'lsSecUnivoca = Mid(Data, 178, 7)
                   'oCCE.CCE_RechazaTRA_Enviada lsSecUnivoca, gdFecSis
                End If
                oCCE.CCE_InsertaIntercambioTransferenciaDet lnIDMN, 0, UCase(Left(Data, 1)), Data, "CHE", 11
            End If
            If lnIDME > 0 And sMoneda = "2" And lnRegME <> 2 Then
                If Left(Data, 1) = "5" Then
                   ' lsTipTra = Mid(Data, 67, 3)
                End If
                If Left(Data, 1) = "6" Then
                   'lsSecUnivoca = Mid(Data, 178, 7)
                   'oCCE.CCE_RechazaTRA_Enviada lsSecUnivoca, gdFecSis
                End If
                oCCE.CCE_InsertaIntercambioTransferenciaDet lnIDME, 0, UCase(Left(Data, 1)), Data, "CHE", 11
            End If
    Wend
    Close #NumeroArchivo
    
     txtResultado.Text = txtResultado.Text & _
                        "Archivo Cargado: " & sNombreArchivo & vbCrLf & _
                        "Registros Cargados: " & (lnRegMN + lnRegME) & vbCrLf
 oCCE.CommitTrans
    lbArch = False
    Exit Sub

ErrorCheP:
    If lbArch Then
        oCCE.RollbackTrans
        Set oCCE = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Private Sub LeerCheque_Rechazados_Recibido(Optional ByRef psFile As String = "")
    Dim oFSO As Scripting.FileSystemObject
    Dim oStream As Scripting.TextStream
    Dim mvarRuta, sNombreArchivo, sRecibido, sId, sFechaArchivo, sCodAplicacion, sNumArchivo, sUsuario, sMoneda, Linea, Data, lsCabDataMN, lsCabDataME As String
    Dim nId, NumeroArchivo As Integer
    Dim lnLoteMN, lnLoteME, lnRegMN, lnRegME As Long
    Dim lbCuentaSoles As Boolean
    Dim lvConfirmado() As String
    Dim lvDevuelto() As String
    Dim lsCabLote, lsCabArchivo, lsRegIndividual, lsRegAdicional, sCCI, sDoi, sVal, sEntAcre, espacios, espacios2, sSecU, sContReg As String
    Dim lnIDMN, lnIDME, lnREF As Long
    Dim lnIDRef As Long
    Dim lsMovNro, lsSecUnivoca, lsTipTra, lsCabNullMN, lsCabNullME As String
    Dim lbArch As Boolean
    Dim sIFICod, lsIfiCodRef As String
    Dim sCtaADeb As String
    Dim sNroCheque, lsNroCheRef As String
    Dim sFechaMB As String
    Dim llIdRef As Long
    Dim lsFechaPreChe As Date
    
On Error GoTo ErrorCheR
oCCE.BeginTrans
lbArch = True
            Set rsConstante = oCCE.CCE_ConstArchivos
            gCCEChequeArch = rsConstante!gCCEChequeArch
            
            lsFechaPreChe = DateAdd("d", -1, gdFecSis)
            
    If Len(psFile) <> 0 Then mvarRuta = psFile
    If Len(mvarRuta) = 0 Then Exit Sub: oCCE.CommitTrans
    
    Set oFSO = New Scripting.FileSystemObject
    Set oStream = oFSO.OpenTextFile(mvarRuta, ForReading)
    sNombreArchivo = oFSO.GetFileName(mvarRuta)
    NumeroArchivo = FreeFile
    'Controles **********************************
    Open mvarRuta For Input As #NumeroArchivo
    While Not EOF(NumeroArchivo)
            Line Input #NumeroArchivo, Linea
            Data = Linea
            Select Case UCase(Left(Data, 1))
                    Case 1
                            If Mid(Data, 3, 1) = 1 Then
                                lbCuentaSoles = True
                                lnLoteMN = 0
                                lnRegMN = 0
                            Else
                                lbCuentaSoles = False
                                lnLoteME = 0
                                lnRegME = 0
                            End If
                    Case 5
                            If lbCuentaSoles Then
                                lnLoteMN = lnLoteMN + 1
                            Else
                                lnLoteME = lnLoteME + 1
                            End If
            End Select
            If lbCuentaSoles Then
                lnRegMN = lnRegMN + 1
            Else
                lnRegME = lnRegME + 1
            End If
    Wend
    Close #NumeroArchivo
    '*******************************************************************
    'Detalle*************
    lnIDMN = 0
    lnIDME = 0
    Open mvarRuta For Input As #NumeroArchivo
     While Not EOF(NumeroArchivo)
            Line Input #NumeroArchivo, Linea
            Data = Linea
            Select Case UCase(Left(Data, 1))
                    Case 1
                            sMoneda = Mid(Data, 3, 1)
                            sFechaArchivo = Mid(Data, 23, 8)
                            sCodAplicacion = Mid(Data, 4, 3)
                            sNumArchivo = CInt(Mid(Data, 31, 2))
                            If sMoneda = "1" Then
                                lsCabDataMN = oCCE.CCE_ValidaCabecera(Data)
                                 lsCabNullMN = oCCE.CCE_ValidaCabeceraNulas(gCCEChequeArch, sFechaArchivo, sMoneda, sCodAplicacion, 12)
                                If lsCabDataMN = "no" And lsCabNullMN = "no" Then
                                lnIDMN = oCCE.CCE_InsIntercambioTransferenciaRec(gCCEChequeArch, sFechaArchivo, sMoneda, sCodAplicacion, sNumArchivo, lnLoteMN, lnRegMN, 1, gsCodUser, 12)
                                Else
                                sFechaMB = Right(sFechaArchivo, 2) & "/" & Mid(sFechaArchivo, 5, 2) & "/" & Left(sFechaArchivo, 4)
                                MsgBox "Ya se realizo la Carga de Rechazados de Cheques del Archivo  " & sNombreArchivo & " ""   del :  " & sFechaMB & " ", vbInformation, "Aviso"
                                Close #NumeroArchivo
                                oCCE.CommitTrans
                                Exit Sub
                                End If
                            Else
                                lsCabDataME = oCCE.CCE_ValidaCabecera(Data)
                                lsCabNullME = oCCE.CCE_ValidaCabeceraNulas(gCCEChequeArch, sFechaArchivo, sMoneda, sCodAplicacion, 12)
                                If lsCabDataME = "no" And lsCabNullME = "no" Then
                                lnIDME = oCCE.CCE_InsIntercambioTransferenciaRec(gCCEChequeArch, sFechaArchivo, sMoneda, sCodAplicacion, sNumArchivo, lnLoteME, lnRegME, 1, gsCodUser, 12)
                                Else
                                sFechaMB = Right(sFechaArchivo, 2) & "/" & Mid(sFechaArchivo, 5, 2) & "/" & Left(sFechaArchivo, 4)
                                MsgBox "Ya se realizo la Carga de Rechazados de Cheques del Archivo  " & sNombreArchivo & " ""   del :  " & sFechaMB & " ", vbInformation, "Aviso"
                                Close #NumeroArchivo
                                oCCE.CommitTrans
                                Exit Sub
                                End If
                            End If
            End Select
            If lnIDMN > 0 And sMoneda = "1" And lnRegMN <> 2 Then
                If Left(Data, 1) = "5" Then
                   ' lsTipTra = Mid(Data, 67, 3)
                End If
                If Left(Data, 1) = "6" Then
                   sIFICod = Mid(Data, 6, 8)
                   lsIfiCodRef = "0" & Mid(Data, 15, 3) & "0" & Mid(Data, 18, 3)
                   sNroCheque = Mid(Data, 48, 9)
                   lsNroCheRef = Mid(Data, 48, 8)
                   sCtaADeb = Mid(Data, 15, 18)
                   
                            oCCE.CCE_Rechaza_CHE gdFecSis, gsCodAge, sIFICod, sCtaADeb, sNroCheque
                            lnREF = oCCE.CCE_ObtieneRefCheque(gdFecSis, lsNroCheRef, lsIfiCodRef, sCtaADeb)
                            
                End If
                oCCE.CCE_InsertaIntercambioTransferenciaDet lnIDMN, IIf(Left(Data, 1) = "6" Or Left(Data, 1) = "7", lnREF, 0), UCase(Left(Data, 1)), Data, "CHE", 12
                
                'lnIDME, IIf(Left(Data, 1) = "6" Or Left(Data, 1) = "7", lsCodSecBCR, 0), IIf(UCase(Left(Data, 1)) = "7", "71", UCase(Left(Data, 1))), Data, IIf(Left(Data, 1) = "5" Or Left(Data, 1) = "6" Or Left(Data, 1) = "7" Or Left(Data, 1) = "8", lsTipTra, 0), 11
            End If
            If lnIDME > 0 And sMoneda = "2" And lnRegME <> 2 Then
                If Left(Data, 1) = "5" Then
                  
                End If
                If Left(Data, 1) = "6" Then
                   sIFICod = Mid(Data, 6, 8)
                   lsIfiCodRef = "0" & Mid(Data, 15, 3) & "0" & Mid(Data, 18, 3)
                   sNroCheque = Mid(Data, 48, 9)
                   lsNroCheRef = Mid(Data, 48, 8)
                   sCtaADeb = Mid(Data, 15, 18)
                            oCCE.CCE_Rechaza_CHE gdFecSis, gsCodAge, sIFICod, sCtaADeb, sNroCheque
                            lnREF = oCCE.CCE_ObtieneRefCheque(gdFecSis, lsNroCheRef, lsIfiCodRef, sCtaADeb)
                End If
                oCCE.CCE_InsertaIntercambioTransferenciaDet lnIDME, IIf(Left(Data, 1) = "6" Or Left(Data, 1) = "7", lnREF, 0), UCase(Left(Data, 1)), Data, "CHE", 12
            End If
    Wend
    Close #NumeroArchivo
    
     txtResultado.Text = txtResultado.Text & _
                        "Archivo Cargado: " & sNombreArchivo & vbCrLf & _
                        "Registros Cargados: " & (lnRegMN + lnRegME) & vbCrLf
 oCCE.CommitTrans
    lbArch = False
    Exit Sub

ErrorCheR:
    If lbArch Then
        oCCE.RollbackTrans
        Set oCCE = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Private Sub LeerCheque_SaldosPrevio_Recibido(Optional ByRef psFile As String = "")
    Dim oFSO As Scripting.FileSystemObject
    Dim oStream As Scripting.TextStream
    Dim mvarRuta, sNombreArchivo, sRecibido, sId, sFechaArchivo, sCodAplicacion, sNumArchivo, sUsuario, sMoneda, Linea, Data, lsCabDataMN, lsCabDataME As String
    Dim nId, NumeroArchivo As Integer
    Dim lnLote, lnReg As Long
    Dim lbCuentaSoles As Boolean
    Dim lvConfirmado() As String
    Dim lvDevuelto() As String
    Dim lsCabLote, lsCabArchivo, lsRegIndividual, lsRegAdicional, sCCI, sDoi, sVal, sEntAcre, espacios, espacios2, sSecU, sContReg As String
    Dim lnIDMN, lnIDME As Long
    Dim lnIDRef As Long
    Dim lsMovNro, lsSecUnivoca, lsCabNullMN As String
    Dim lbArch As Boolean
    Dim sFechaMB As String

On Error GoTo ErrorCheSP
oCCE.BeginTrans
lbArch = True
        
    Set rsConstante = oCCE.CCE_ConstArchivos
    gCCESalChequeArch = rsConstante!gCCESalChequeArch
        
    If Len(psFile) <> 0 Then mvarRuta = psFile
    If Len(mvarRuta) = 0 Then Exit Sub: oCCE.CommitTrans
    
    Set oFSO = New Scripting.FileSystemObject
    Set oStream = oFSO.OpenTextFile(mvarRuta, ForReading)
    sNombreArchivo = oFSO.GetFileName(mvarRuta)
    NumeroArchivo = FreeFile
    'Controles **********************************
    Open mvarRuta For Input As #NumeroArchivo
    While Not EOF(NumeroArchivo)
            Line Input #NumeroArchivo, Linea
            Data = Linea
            Select Case UCase(Left(Data, 2))
                    Case "CA"
                                lnLote = 0
                                lnReg = 0
                    Case "CL"
                            sMoneda = Mid(Data, 7, 1)
                            lnLote = lnLote + 1
            End Select
               lnReg = lnReg + 1
            
    Wend
    Close #NumeroArchivo
    '***********************************************
    'Detalle***************
    lnIDMN = 0
    lnIDME = 0
    Open mvarRuta For Input As #NumeroArchivo
     While Not EOF(NumeroArchivo)
            Line Input #NumeroArchivo, Linea
            Data = Linea
            Select Case UCase(Left(Data, 2))
                    Case "CA"
                            sFechaArchivo = Mid(Data, 43, 8)
                            sCodAplicacion = Mid(Data, 19, 3)
                            sNumArchivo = CInt(Mid(Data, 41, 1))
                            lsCabDataMN = oCCE.CCE_ValidaCabecera(Data)
                             lsCabNullMN = oCCE.CCE_ValidaCabeceraNulas(gCCESalChequeArch, sFechaArchivo, sMoneda, sCodAplicacion, 15)
                                If lsCabDataMN = "no" And lsCabNullMN = "no" Then
                                lnIDMN = oCCE.CCE_InsIntercambioTransferenciaRec(gCCESalChequeArch, sFechaArchivo, sMoneda, sCodAplicacion, sNumArchivo, lnLote, lnReg, 1, gsCodUser, 15)
                                Else
                                sFechaMB = Right(sFechaArchivo, 2) & "/" & Mid(sFechaArchivo, 5, 2) & "/" & Left(sFechaArchivo, 4)
                                MsgBox "Ya se realizo la Carga de Saldos Previos de Cheques del Archivo  " & sNombreArchivo & " ""   del :  " & sFechaMB & " ", vbInformation, "Aviso"
                                Close #NumeroArchivo
                                oCCE.CommitTrans
                                Exit Sub
                                End If
         End Select
            If lnIDMN > 0 And lnReg <> 2 Then
                oCCE.CCE_InsertaIntercambioTransferenciaDet lnIDMN, 0, UCase(Left(Data, 2)), Data, "CHE", 15
            End If
    Wend
    Close #NumeroArchivo

       txtResultado.Text = txtResultado.Text & _
                        "Archivo Cargado: " & sNombreArchivo & vbCrLf & _
                        "Registros Cargados: " & lnReg & vbCrLf
 oCCE.CommitTrans
    lbArch = False
    Exit Sub

ErrorCheSP:
    If lbArch Then
        oCCE.RollbackTrans
        Set oCCE = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Private Sub LeerCheque_SaldosRCG_Recibido(Optional ByRef psFile As String = "")
    Dim oFSO As Scripting.FileSystemObject
    Dim oStream As Scripting.TextStream
    Dim mvarRuta, sNombreArchivo, sRecibido, sId, sFechaArchivo, sCodAplicacion, sNumArchivo, sUsuario, sMoneda, Linea, Data, lsCabDataMN, lsCabDataME As String
    Dim nId, NumeroArchivo As Integer
    Dim lnLote, lnReg As Long
    Dim lbCuentaSoles As Boolean
    Dim lvConfirmado() As String
    Dim lvDevuelto() As String
    Dim lsCabLote, lsCabArchivo, lsRegIndividual, lsRegAdicional, sCCI, sDoi, sVal, sEntAcre, espacios, espacios2, sSecU, sContReg As String
    Dim lnIDMN, lnIDME As Long
    Dim lnIDRef As Long
    Dim lsMovNro, lsSecUnivoca, lsCabNullMN As String
    Dim lbArch As Boolean
    Dim sFechaMB As String

On Error GoTo ErrorCheRCG
oCCE.BeginTrans
lbArch = True
        
    Set rsConstante = oCCE.CCE_ConstArchivos
    gCCESalChequeArch = rsConstante!gCCESalChequeArch
        
    If Len(psFile) <> 0 Then mvarRuta = psFile
    If Len(mvarRuta) = 0 Then Exit Sub: oCCE.CommitTrans
    
    Set oFSO = New Scripting.FileSystemObject
    Set oStream = oFSO.OpenTextFile(mvarRuta, ForReading)
    sNombreArchivo = oFSO.GetFileName(mvarRuta)
    NumeroArchivo = FreeFile
    'Controles **********************************
    Open mvarRuta For Input As #NumeroArchivo
    While Not EOF(NumeroArchivo)
            Line Input #NumeroArchivo, Linea
            Data = Linea
            Select Case UCase(Left(Data, 2))
                    Case "CA"
                                lnLote = 0
                                lnReg = 0
                    Case "CL"
                            sMoneda = Mid(Data, 7, 1)
                            lnLote = lnLote + 1
            End Select
               lnReg = lnReg + 1
            
    Wend
    Close #NumeroArchivo
    '***********************************************
    'Detalle***************
    lnIDMN = 0
    lnIDME = 0
    Open mvarRuta For Input As #NumeroArchivo
     While Not EOF(NumeroArchivo)
            Line Input #NumeroArchivo, Linea
            Data = Linea
            Select Case UCase(Left(Data, 2))
                    Case "CA"
                            sFechaArchivo = Mid(Data, 43, 8)
                            sCodAplicacion = Mid(Data, 19, 3)
                            sNumArchivo = CInt(Mid(Data, 41, 1))
                            lsCabDataMN = oCCE.CCE_ValidaCabecera(Data)
                             lsCabNullMN = oCCE.CCE_ValidaCabeceraNulas(gCCESalChequeArch, sFechaArchivo, sMoneda, sCodAplicacion, 19)
                                If lsCabDataMN = "no" And lsCabNullMN = "no" Then
                                lnIDMN = oCCE.CCE_InsIntercambioTransferenciaRec(gCCESalChequeArch, sFechaArchivo, sMoneda, sCodAplicacion, sNumArchivo, lnLote, lnReg, 1, gsCodUser, 19)
                                Else
                                 sFechaMB = Right(sFechaArchivo, 2) & "/" & Mid(sFechaArchivo, 5, 2) & "/" & Left(sFechaArchivo, 4)
                                MsgBox "Ya se realizo la Carga de Saldos de RCG de Cheques del Archivo  " & sNombreArchivo & " ""   del :  " & sFechaMB & " ", vbInformation, "Aviso"
                                Close #NumeroArchivo
                                oCCE.CommitTrans
                                Exit Sub
                                End If
         End Select
            If lnIDMN > 0 And lnReg <> 2 Then
                oCCE.CCE_InsertaIntercambioTransferenciaDet lnIDMN, 0, UCase(Left(Data, 2)), Data, "CHE", 19
            End If
    Wend
    Close #NumeroArchivo

       txtResultado.Text = txtResultado.Text & _
                        "Archivo Cargado: " & sNombreArchivo & vbCrLf & _
                        "Registros Cargados: " & lnReg & vbCrLf
 oCCE.CommitTrans
    lbArch = False
    Exit Sub

ErrorCheRCG:
    If lbArch Then
        oCCE.RollbackTrans
        Set oCCE = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Private Sub LeerCheque_SaldosFinales_Recibido(Optional ByRef psFile As String = "")
    Dim oFSO As Scripting.FileSystemObject
    Dim oStream As Scripting.TextStream
    Dim mvarRuta, sNombreArchivo, sRecibido, sId, sFechaArchivo, sCodAplicacion, sNumArchivo, sUsuario, sMoneda, Linea, Data, lsCabDataMN, lsCabDataME As String
    Dim nId, NumeroArchivo As Integer
    Dim lnLote, lnReg As Long
    Dim lbCuentaSoles As Boolean
    Dim lvConfirmado() As String
    Dim lvDevuelto() As String
    Dim lsCabLote, lsCabArchivo, lsRegIndividual, lsRegAdicional, sCCI, sDoi, sVal, sEntAcre, espacios, espacios2, sSecU, sContReg As String
    Dim lnIDMN, lnIDME As Long
    Dim lnIDRef As Long
    Dim lsMovNro, lsSecUnivoca, lsCabNullMN As String
    Dim lbArch As Boolean
    Dim sFechaMB As String
    
On Error GoTo ErrorCheSF
oCCE.BeginTrans
lbArch = True
        
    Set rsConstante = oCCE.CCE_ConstArchivos
    gCCESalChequeArch = rsConstante!gCCESalChequeArch
        
    If Len(psFile) <> 0 Then mvarRuta = psFile
    If Len(mvarRuta) = 0 Then Exit Sub: oCCE.CommitTrans
    
    Set oFSO = New Scripting.FileSystemObject
    Set oStream = oFSO.OpenTextFile(mvarRuta, ForReading)
    sNombreArchivo = oFSO.GetFileName(mvarRuta)
    NumeroArchivo = FreeFile
    'Controles **********************************
    Open mvarRuta For Input As #NumeroArchivo
    While Not EOF(NumeroArchivo)
            Line Input #NumeroArchivo, Linea
            Data = Linea
            Select Case UCase(Left(Data, 2))
                    Case "CA"
                                lnLote = 0
                                lnReg = 0
                    Case "CL"
                            sMoneda = Mid(Data, 7, 1)
                            lnLote = lnLote + 1
            End Select
               lnReg = lnReg + 1
            
    Wend
    Close #NumeroArchivo
    '***********************************************
    'Detalle***************
    lnIDMN = 0
    lnIDME = 0
    Open mvarRuta For Input As #NumeroArchivo
     While Not EOF(NumeroArchivo)
            Line Input #NumeroArchivo, Linea
            Data = Linea
            Select Case UCase(Left(Data, 2))
                    Case "CA"
                            sFechaArchivo = Mid(Data, 43, 8)
                            sCodAplicacion = Mid(Data, 19, 3)
                            sNumArchivo = CInt(Mid(Data, 41, 1))
                            lsCabDataMN = oCCE.CCE_ValidaCabecera(Data)
                             lsCabNullMN = oCCE.CCE_ValidaCabeceraNulas(gCCESalChequeArch, sFechaArchivo, sMoneda, sCodAplicacion, 16)
                                If lsCabDataMN = "no" And lsCabNullMN = "no" Then
                                lnIDMN = oCCE.CCE_InsIntercambioTransferenciaRec(gCCESalChequeArch, sFechaArchivo, sMoneda, sCodAplicacion, sNumArchivo, lnLote, lnReg, 1, gsCodUser, 16)
                                Else
                                 sFechaMB = Right(sFechaArchivo, 2) & "/" & Mid(sFechaArchivo, 5, 2) & "/" & Left(sFechaArchivo, 4)
                                MsgBox "Ya se realizo la Carga de Saldos de RCG de Cheques del Archivo  " & sNombreArchivo & " ""   del :  " & sFechaMB & " ", vbInformation, "Aviso"
                                Close #NumeroArchivo
                                oCCE.CommitTrans
                                Exit Sub
                                End If
         End Select
            If lnIDMN > 0 And lnReg <> 2 Then
                oCCE.CCE_InsertaIntercambioTransferenciaDet lnIDMN, 0, UCase(Left(Data, 2)), Data, "CHE", 16
            End If
    Wend
    Close #NumeroArchivo

       txtResultado.Text = txtResultado.Text & _
                        "Archivo Cargado: " & sNombreArchivo & vbCrLf & _
                        "Registros Cargados: " & lnReg & vbCrLf
 oCCE.CommitTrans
    lbArch = False
    Exit Sub

ErrorCheSF:
    If lbArch Then
        oCCE.RollbackTrans
        Set oCCE = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Private Sub LeerCheque_Nulo_Liquidacion(Optional ByRef psFile As String = "")
 Dim oFSO As Scripting.FileSystemObject
    Dim oStream As Scripting.TextStream
    Dim mvarRuta, sNombreArchivo, sRecibido, sId, sFechaArchivo, sCodAplicacion, sNumArchivo, sUsuario, sMoneda, Linea, Data, lsCabDataMN, lsCabDataME As String
    Dim nId, NumeroArchivo As Integer
    Dim lnLote, lnReg As Long
    Dim lbCuentaSoles As Boolean
    Dim lvConfirmado() As String
    Dim lvDevuelto() As String
    Dim lsCabLote, lsCabArchivo, lsRegIndividual, lsRegAdicional, sCCI, sDoi, sVal, sEntAcre, espacios, espacios2, sSecU, sContReg As String
    Dim lnIDMN, lnIDME As Long
    Dim lnIDRef As Long
    Dim lsMovNro, lsSecUnivoca, lsCabNullMN As String
    Dim lbArch As Boolean
    Dim sFechaMB As String

On Error GoTo ErrorCheNL
oCCE.BeginTrans
lbArch = True
        
    Set rsConstante = oCCE.CCE_ConstArchivos
    gCCEConfLiqCheque = rsConstante!gCCEConfLiqCheque
    
        
    If Len(psFile) <> 0 Then mvarRuta = psFile
    If Len(mvarRuta) = 0 Then Exit Sub: oCCE.CommitTrans
    
    Set oFSO = New Scripting.FileSystemObject
    Set oStream = oFSO.OpenTextFile(mvarRuta, ForReading)
    sNombreArchivo = oFSO.GetFileName(mvarRuta)
    NumeroArchivo = FreeFile
    'Controles **********************************
    Open mvarRuta For Input As #NumeroArchivo
    While Not EOF(NumeroArchivo)
            Line Input #NumeroArchivo, Linea
            Data = Linea
            Select Case UCase(Left(Data, 2))
                    Case "CA"
                                lnLote = 0
                                lnReg = 0
                    Case "CL"
                            lnLote = lnLote + 1
            End Select
               lnReg = lnReg + 1
            
    Wend
    Close #NumeroArchivo
    '***********************************************
    'Detalle*************
    lnIDMN = 0
    lnIDME = 0
    Open mvarRuta For Input As #NumeroArchivo
     While Not EOF(NumeroArchivo)
            Line Input #NumeroArchivo, Linea
            Data = Linea
            Select Case UCase(Left(Data, 2))
                    Case "CA"
                            sFechaArchivo = Mid(Data, 43, 8)
                            sMoneda = Mid(Data, 51, 1)
                            sCodAplicacion = Mid(Data, 19, 3)
                            sNumArchivo = CInt(Mid(Data, 41, 1))
                            lsCabDataMN = oCCE.CCE_ValidaCabecera(Data)
                             lsCabNullMN = oCCE.CCE_ValidaCabeceraNulas(gCCEConfLiqCheque, sFechaArchivo, sMoneda, sCodAplicacion, 17)
                                If lsCabDataMN = "no" And lsCabNullMN = "no" Then
                                lnIDMN = oCCE.CCE_InsIntercambioTransferenciaRec(gCCEConfLiqCheque, sFechaArchivo, sMoneda, sCodAplicacion, sNumArchivo, lnLote, lnReg, 1, gsCodUser, 17)
                                Else
                                 sFechaMB = Right(sFechaArchivo, 2) & "/" & Mid(sFechaArchivo, 5, 2) & "/" & Left(sFechaArchivo, 4)
                                MsgBox "Ya se realizo la Carga de Archivo de Nulo de Liquidacion de  Cheques del Archivo  " & sNombreArchivo & " ""   del :  " & sFechaMB & " ", vbInformation, "Aviso"
                                Close #NumeroArchivo
                                oCCE.CommitTrans
                                Exit Sub
                                End If
         End Select
            If lnIDMN > 0 Then
                oCCE.CCE_InsertaIntercambioTransferenciaDet lnIDMN, 0, UCase(Left(Data, 2)), Data, "CHE", 17
            End If
    Wend
    Close #NumeroArchivo

       txtResultado.Text = txtResultado.Text & _
                        "Archivo Cargado: " & sNombreArchivo & vbCrLf & _
                        "Registros Cargados: " & lnReg & vbCrLf
 oCCE.CommitTrans
    lbArch = False
    Exit Sub

ErrorCheNL:
    If lbArch Then
        oCCE.RollbackTrans
        Set oCCE = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub cmdCargar_Click()
  LeerTramaRecibido
  Dim i As Integer
PbCCE.Max = 10
PbCCE.value = 0
    For i = 0 To PbCCE.Max
        PbCCE.value = i
        lblPB = CLng((PbCCE.value * 100) / PbCCE.Max) & " %"
        DoEvents
    Next
    Call CargarHorarios
End Sub
Private Sub cmdLimpiar_Click()
    PbCCE.value = 0
    lblPB = 0
    txtResultado.Refresh
    txtResultado.Text = ""
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub CargarHorarios()
Dim rsHorario As ADODB.Recordset
Dim rsHorarioChe As ADODB.Recordset
Set oCamara = New COMNCajaGeneral.NCOMCCE
Set rsHorario = oCamara.CCE_HorarioTransferencia
Set rsHorarioChe = oCamara.CCE_HorarioCheque
 If Not (rsHorario.EOF And rsHorario.BOF) Then
 lblTra.Caption = rsHorario!cCodAplicacion
lblTraDes.Caption = rsHorario!cDesCodAplicacion
lblIniTra.Caption = rsHorario!dHoraTransArchIni
lblFinTra.Caption = rsHorario!dHoraTransArchFin
Else
lblTra.BackColor = &HFF&
lblTraDes.BackColor = &HFF&
lblIniTra.Caption = ""
lblFinTra.Caption = ""
 End If
 If Not (rsHorarioChe.EOF And rsHorarioChe.BOF) Then
 lblChe.Caption = rsHorarioChe!cCodAplicacion
lblCheDes.Caption = rsHorarioChe!cDesSesion
lblIniChe.Caption = rsHorarioChe!dHoraTransArchIni
lblFinChe.Caption = rsHorarioChe!dHoraTransArchFin
Else
lblChe.BackColor = &HFF&
lblCheDes.BackColor = &HFF&
lblIniChe.Caption = ""
lblFinChe.Caption = ""
 End If
End Sub
Private Sub Command1_Click()
Call CargarHorarios
End Sub
Private Sub Form_Load()
Dim rsCarpeta As ADODB.Recordset
Dim i, sTexto
        txtResultado.Text = ""
            Set oCCE = New COMNCajaGeneral.NCOMCCE
            Set oCCER = New COMNCajaGeneral.NCOMCCE
            Set rsCarpeta = oCCE.CCE_ObtieneArchivosBlanco
            txtCarpeta = rsCarpeta!nConsSisValor
            Call CargarHorarios
End Sub
