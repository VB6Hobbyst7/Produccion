VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredFormEvalCredCel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingresos y Egresos MN y ME"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   Icon            =   "frmCredFormEvalCredCel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab stMultipleOp 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4471
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Evaluación"
      TabPicture(0)   =   "frmCredFormEvalCredCel.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frIngEgr"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdCancelarFE"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdGuardarFE"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.CommandButton cmdGuardarFE 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   2520
         TabIndex        =   15
         Top             =   2040
         Width           =   855
      End
      Begin VB.CommandButton cmdCancelarFE 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   3480
         TabIndex        =   14
         Top             =   2040
         Width           =   855
      End
      Begin VB.Frame frIngEgr 
         Caption         =   "Ingresos y Egresos"
         Height          =   1335
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   4215
         Begin SICMACT.EditMoney edmEgrME 
            Height          =   255
            Left            =   3000
            TabIndex        =   2
            Top             =   720
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0"
            Enabled         =   -1  'True
         End
         Begin SICMACT.EditMoney edmEgrMN 
            Height          =   255
            Left            =   3000
            TabIndex        =   3
            Top             =   360
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0"
            Enabled         =   -1  'True
         End
         Begin SICMACT.EditMoney edmIngrME 
            Height          =   255
            Left            =   970
            TabIndex        =   4
            Top             =   720
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0"
            Enabled         =   -1  'True
         End
         Begin SICMACT.EditMoney edmIngrMN 
            Height          =   255
            Left            =   970
            TabIndex        =   5
            Top             =   360
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0"
            Enabled         =   -1  'True
         End
         Begin VB.Label Label13 
            Caption         =   "%"
            Height          =   255
            Left            =   1850
            TabIndex        =   13
            Top             =   765
            Width           =   255
         End
         Begin VB.Label Label12 
            Caption         =   "%"
            Height          =   255
            Left            =   1850
            TabIndex        =   12
            Top             =   405
            Width           =   255
         End
         Begin VB.Label Label11 
            Caption         =   "%"
            Height          =   255
            Left            =   3880
            TabIndex        =   11
            Top             =   765
            Width           =   255
         End
         Begin VB.Label Label10 
            Caption         =   "%"
            Height          =   255
            Left            =   3880
            TabIndex        =   10
            Top             =   405
            Width           =   255
         End
         Begin VB.Label Label9 
            Caption         =   "Egreso ME"
            Height          =   255
            Left            =   2160
            TabIndex        =   9
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label8 
            Caption         =   "Egreso MN"
            Height          =   255
            Left            =   2160
            TabIndex        =   8
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label7 
            Caption         =   "Ingreso ME"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label6 
            Caption         =   "Ingreso MN"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   975
         End
      End
   End
End
Attribute VB_Name = "frmCredFormEvalCredCel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================
'Autor: JOEP
'Fecha: 20190124
'ERS034: Riesgo Crediticios JOEP - RUSI
'===============================================================
Option Explicit
Dim cCtaCod As String
Dim nTpReg As Integer
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Const MF_BYPOSITION = &H400&

Public Sub Inicio(ByVal pcCtaCod As String, Optional ByVal pnTpReg As Integer)
    cCtaCod = ""
    nTpReg = pnTpReg
If (pnTpReg = 10 Or pnTpReg = 11) Then
    cCtaCod = pcCtaCod
    Caption = "Ingresos y Egresos MN y ME"
    Call CargaDatos
End If
    Show 1
End Sub

'tab 1
Private Sub edmEgrME_LostFocus()
    edmEgrME = IIf(Len(edmEgrME) = 1, IIf(edmEgrME = ".", 0, edmEgrME), IIf(Len(edmEgrME) = 2 And CDbl(Replace(edmEgrME, ".", 0)) = 0, 0, edmEgrME))
End Sub

Private Sub edmEgrME_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    edmEgrME = IIf(Len(edmEgrME) = 1, IIf(edmEgrME = ".", 0, edmEgrME), IIf(Len(edmEgrME) = 2 And CDbl(Replace(edmEgrME, ".", 0)) = 0, 0, edmEgrME))
        If Not ValidaMonto(edmEgrME) Then
            EnfocaControl cmdGuardarFE
        Else
            edmEgrME = 0
        End If
    End If
End Sub

Private Sub edmEgrMN_LostFocus()
    edmEgrMN = IIf(Len(edmEgrMN) = 1, IIf(edmEgrMN = ".", 0, edmEgrMN), IIf(Len(edmEgrMN) = 2 And CDbl(Replace(edmEgrMN, ".", 0)) = 0, 0, edmEgrMN))
End Sub

Private Sub edmEgrMN_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    edmEgrMN = IIf(Len(edmEgrMN) = 1, IIf(edmEgrMN = ".", 0, edmEgrMN), IIf(Len(edmEgrMN) = 2 And CDbl(Replace(edmEgrMN, ".", 0)) = 0, 0, edmEgrMN))
        If Not ValidaMonto(edmEgrMN) Then
            EnfocaControl edmIngrME
        Else
            edmEgrMN = 0
        End If
    End If
End Sub

Private Sub edmIngrME_LostFocus()
    edmIngrME = IIf(Len(edmIngrME) = 1, IIf(edmIngrME = ".", 0, edmIngrME), IIf(Len(edmIngrME) = 2 And CDbl(Replace(edmIngrME, ".", 0)) = 0, 0, edmIngrME))
End Sub

Private Sub edmIngrME_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    edmIngrME = IIf(Len(edmIngrME) = 1, IIf(edmIngrME = ".", 0, edmIngrME), IIf(Len(edmIngrME) = 2 And CDbl(Replace(edmIngrME, ".", 0)) = 0, 0, edmIngrME))
        If Not ValidaMonto(edmIngrME) Then
           EnfocaControl edmEgrME
        Else
            edmIngrME = 0
        End If
    End If
End Sub

Private Function ValidaMonto(ByVal pnMonto As Double) As Boolean
    ValidaMonto = False
    If pnMonto > 100 Then
        MsgBox "Solo se acepta montos en el rango [0 - 100]%", vbInformation, "Aviso"
        ValidaMonto = True
    End If
End Function

Public Function DisableCloseButton(frm As Form) As Boolean
'PURPOSE: Removes X button from a form
'EXAMPLE: DisableCloseButton Me
'RETURNS: True if successful, false otherwise
'NOTES:   Also removes Exit Item from
'         Control Box Menu
    Dim lHndSysMenu As Long
    Dim lAns1 As Long, lAns2 As Long
    lHndSysMenu = GetSystemMenu(frm.hwnd, 0)
    'remove close button
    lAns1 = RemoveMenu(lHndSysMenu, 6, MF_BYPOSITION)
   'Remove seperator bar
    lAns2 = RemoveMenu(lHndSysMenu, 5, MF_BYPOSITION)
    'Return True if both calls were successful
    DisableCloseButton = (lAns1 <> 0 And lAns2 <> 0)
End Function
'tab 1

Private Sub Form_Load()
    CentraForm Me
End Sub

Private Sub edmIngrMN_LostFocus()
    edmIngrMN = IIf(Len(edmIngrMN) = 1, IIf(edmIngrMN = ".", 0, edmIngrMN), IIf(Len(edmIngrMN) = 2 And CDbl(Replace(edmIngrMN, ".", 0)) = 0, 0, edmIngrMN))
End Sub

Private Sub edmIngrMN_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    edmIngrMN = IIf(Len(edmIngrMN) = 1, IIf(edmIngrMN = ".", 0, edmIngrMN), IIf(Len(edmIngrMN) = 2 And CDbl(Replace(edmIngrMN, ".", 0)) = 0, 0, edmIngrMN))
        If Not ValidaMonto(edmIngrMN) Then
            EnfocaControl edmEgrMN
        Else
            edmIngrMN = 0
        End If
    End If
End Sub

Private Sub cmdGuardarFE_Click()
Dim rsGrabaIngEgr As ADODB.Recordset
Dim obCredIngEgr As COMDCredito.DCOMCredito
Set obCredIngEgr = New COMDCredito.DCOMCredito

If Not ValidadIngEgr Then
    Set rsGrabaIngEgr = obCredIngEgr.GrabaRigCamCred(cCtaCod, CDbl(edmIngrMN), CDbl(edmEgrMN), CDbl(edmIngrME), CDbl(edmEgrME))
    
    If rsGrabaIngEgr!nResl = 1 Then
        Unload Me
    Else
        MsgBox "Error, no se registraron los datos correctamente de Ingresos y Egresos - Comuniquese con TI-Desarrollo", vbInformation, "Aviso"
    End If
End If

RSClose rsGrabaIngEgr
Set obCredIngEgr = Nothing
End Sub

Private Function ValidadIngEgr() As Boolean
ValidadIngEgr = False
If (CDbl(edmIngrMN.Text) + CDbl(edmIngrME.Text)) <> 100 Then
    MsgBox "Los Ingresos MN y Ingresos en ME tiene que sumar 100%", vbInformation, "Aviso"
    ValidadIngEgr = True
    Exit Function
End If
If (CDbl(edmEgrMN.Text) + CDbl(edmEgrME.Text)) <> 100 Then
    MsgBox "Los Egresos MN y Egresos en ME tiene que sumar 100%", vbInformation, "Aviso"
    ValidadIngEgr = True
    Exit Function
End If
End Function

Private Sub cmdCancelarFE_Click()
    Unload Me
End Sub

Private Sub habilitarControlesCrediCel()
    'tab 1
    If nTpReg = 10 Then
        cmdCancelarFE.Enabled = False
        DisableCloseButton Me
    End If
    If nTpReg = 11 Then
         cmdGuardarFE.Enabled = False
         edmIngrMN.Enabled = False
         edmEgrMN.Enabled = False
         
         edmIngrME.Enabled = False
         edmEgrME.Enabled = False
    End If
    'tab 1
End Sub

Private Sub CargaDatos()
'tab 1
If (nTpReg = 10 Or nTpReg = 11) Then
    Dim obCredIG As COMDCredito.DCOMFormatosEval
    Dim rsCargaDatosIG As ADODB.Recordset
    Set obCredIG = New COMDCredito.DCOMFormatosEval
    Set rsCargaDatosIG = obCredIG.RecuperaDatosRatios(cCtaCod)
    If Not (rsCargaDatosIG.BOF And rsCargaDatosIG.EOF) Then
        edmIngrMN = Format(rsCargaDatosIG!nIngMN, "#,#0.00")
        edmEgrMN = Format(rsCargaDatosIG!nEgrMN, "#,#0.00")
        edmIngrME = Format(rsCargaDatosIG!nIngME, "#,#0.00")
        edmEgrME = Format(rsCargaDatosIG!nEgrME, "#,#0.00")
    End If
    RSClose rsCargaDatosIG
    Set obCredIG = Nothing
    Call habilitarControlesCrediCel
End If
'tab 1
End Sub
