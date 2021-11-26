VERSION 5.00
Begin VB.Form frmSetupCOM 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4305
   Icon            =   "frmSetupCOM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "PenWare"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   3855
      Begin VB.ComboBox cboTipoPenware 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   360
         Width           =   2535
      End
      Begin VB.ComboBox cboPenware 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   420
         Width           =   315
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Puerto"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   900
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Aceptar"
      Height          =   400
      Left            =   960
      TabIndex        =   4
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   400
      Left            =   2520
      TabIndex        =   3
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Frame fraDispositivo 
      Caption         =   "PinPad"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.ComboBox cboTipoPinPad 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   360
         Width           =   2535
      End
      Begin VB.ComboBox cboPinPad 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   900
         Width           =   2535
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   420
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Puerto"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   960
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmSetupCOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpbuffer As String, nSize As Long) As Long
Dim sMaquina As String
Dim buffMaq As String
Dim lSizeMaq As Long

Private Sub IniciaCombo(ByRef cboConst As ComboBox, ByVal nCapConst As ConstanteCabecera)
Dim clsGen As COMDConstSistema.DCOMGeneral
Dim rsConst As ADODB.Recordset
Set clsGen = New COMDConstSistema.DCOMGeneral
Set rsConst = clsGen.GetConstante(nCapConst)
Set clsGen = Nothing
Do While Not rsConst.EOF
    cboConst.AddItem rsConst("cDescripcion") & Space(100) & rsConst("nConsValor")
    rsConst.MoveNext
Loop
cboConst.ListIndex = 0
End Sub

Private Sub SetupActual()
Dim clsGen As COMDConstSistema.DCOMGeneral
Dim rsPC As ADODB.Recordset
Dim nPeriferico As TipoPeriferico
Dim nPuerto As TipoPuertoSerial

On Error GoTo ErrSetup

'Verifica la configuración inicial de la máquina
cboPinPad.ListIndex = 0
cboPenware.ListIndex = 0

Set clsGen = New COMDConstSistema.DCOMGeneral
Set rsPC = clsGen.GetPerifericosPC(sMaquina)

UbicaCombo cboTipoPinPad, rsPC("nMarca"), , 4

If rsPC.BOF And rsPC.EOF Then
    cboPinPad.ListIndex = 0
    cboPenware.ListIndex = 0
Else
    Do While Not rsPC.EOF
        nPeriferico = rsPC("nPeriferico")
        nPuerto = rsPC("nPuerto")
        If nPeriferico = gPerifPINPAD Then
            cboPinPad.ListIndex = nPuerto
        ElseIf nPeriferico = gPerifPENWARE Then
            cboPenware.ListIndex = nPuerto
        End If
        rsPC.MoveNext
    Loop
End If
rsPC.Close
Set rsPC = Nothing
Exit Sub
ErrSetup:
    MsgBox Err.Description, vbExclamation, "Error del Sistema"
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()

If (cboPinPad.ListIndex = cboPenware.ListIndex) And cboPinPad.ListIndex > 0 Then
    MsgBox "No es posible establecer dos dispositivos en un mismo puerto serial", vbExclamation, "Aviso"
Else

    If MsgBox("Desea grabar la información", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    
        Dim clsGen As COMDConstSistema.DCOMGeneral
        Dim nPuerto As TipoPuertoSerial
        
        Dim nTipoPinPad  As TipoPinPad
        Dim nTipoPenware As Integer    'Crear su Tipo Si se maneje Penware
        
        
        
        Set clsGen = New COMDConstSistema.DCOMGeneral
        
        clsGen.EliminaPerifericosPC sMaquina
        
        If cboPinPad.ListIndex > 0 Then
            nPuerto = CLng(Trim(Right(cboPinPad.Text, 5)))
            nTipoPinPad = CLng(Trim(Right(cboTipoPinPad.Text, 5)))
            clsGen.AgergaPerifericoPC sMaquina, gPerifPINPAD, nPuerto, nTipoPinPad
        End If
        If cboPenware.ListIndex > 0 Then
            nPuerto = CLng(Trim(Right(cboPenware.Text, 5)))
            nTipoPenware = 0  'CLng(Trim(Right(cboTipoPenware.Text, 5)))
            clsGen.AgergaPerifericoPC sMaquina, gPerifPENWARE, nPuerto, nTipoPenware
        End If
        Unload Me
    End If
End If
End Sub

Private Sub Form_Load()
Me.Caption = "Configuración Periféricos"
'Obtiene el nombre de la PC
Me.Icon = LoadPicture(App.path & gsRutaIcono)

buffMaq = Space(255)
lSizeMaq = Len(buffMaq)
Call GetComputerName(buffMaq, lSizeMaq)
sMaquina = Trim(Left$(buffMaq, lSizeMaq))
cboPinPad.AddItem "<Ninguno>"
IniciaCombo cboPinPad, gTipoPuertoSerial
cboPenware.AddItem "<Ninguno>"
IniciaCombo cboPenware, gTipoPuertoSerial
IniciaCombo cboTipoPinPad, gTipoPinPad
SetupActual

End Sub

