VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredMontoPresupuestado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Monto Presupuestado"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3165
   Icon            =   "frmCredMontoPresupuestado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   3165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab stDatos 
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   4895
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      ForeColor       =   -2147483635
      TabCaption(0)   =   "Datos"
      TabPicture(0)   =   "frmCredMontoPresupuestado.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frDatos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame frDatos 
         Caption         =   "Datos"
         Height          =   2175
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   2895
         Begin VB.CommandButton cmdGuardar 
            Caption         =   "Guardar"
            Height          =   255
            Left            =   1920
            TabIndex        =   13
            Top             =   1800
            Width           =   855
         End
         Begin SICMACT.EditMoney EditMPorcentaje 
            Height          =   255
            Left            =   1800
            TabIndex        =   12
            Top             =   720
            Width           =   975
            _ExtentX        =   1720
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
         Begin SICMACT.EditMoney editMMontoPresu 
            Height          =   255
            Left            =   1440
            TabIndex        =   11
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
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
         Begin SICMACT.EditMoney editMMontoSoli 
            Height          =   255
            Left            =   1440
            TabIndex        =   15
            Top             =   1440
            Width           =   1335
            _ExtentX        =   2355
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
         Begin SICMACT.EditMoney editMMontoDisp 
            Height          =   255
            Left            =   1440
            TabIndex        =   16
            Top             =   1080
            Width           =   1335
            _ExtentX        =   2355
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
         End
         Begin VB.Label lblMontoDisp 
            Caption         =   "Monto Disponible"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label lblMontSol 
            Caption         =   "Monto Solicitado"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "%:"
            Height          =   255
            Left            =   1440
            TabIndex        =   10
            Top             =   720
            Width           =   255
         End
         Begin VB.Label lblDesc 
            Caption         =   "Monto:"
            Height          =   495
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " Ventas y Costos "
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
         Height          =   735
         Left            =   0
         TabIndex        =   1
         Top             =   -2640
         Width           =   9975
         Begin SICMACT.EditMoney txtIngresoNegocio 
            Height          =   300
            Left            =   1800
            TabIndex        =   2
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
         Begin SICMACT.EditMoney txtMargenBruto 
            Height          =   300
            Left            =   8640
            TabIndex        =   3
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
         Begin SICMACT.EditMoney txtEgresoNegocio 
            Height          =   300
            Left            =   4800
            TabIndex        =   4
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ingresos del Negocio :"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   285
            Width           =   1605
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Egreso por Venta:"
            Height          =   195
            Left            =   3360
            TabIndex        =   6
            Top             =   285
            Width           =   1305
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Margen Bruto :"
            Height          =   195
            Left            =   7440
            TabIndex        =   5
            Top             =   285
            Width           =   1080
         End
      End
   End
End
Attribute VB_Name = "frmCredMontoPresupuestado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'Autor: JOEP
'ERS 042 - Catalogo de Productos
'Fecha: 2019-02-08
'*****************************************************************************************************
Option Explicit

Dim nPorcentaje As Double
Dim nMatMontoSol As Variant
Dim cTpProd As String
Dim nMontMin As Currency
Dim nTpMonedad As Integer

Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Const MF_BYPOSITION = &H400&

Public Sub Inicio(ByVal pnPorcentaje As Double, ByRef pnMatMontoSol As Variant, ByVal nDestino As Double, Optional ByVal pcTpProd As String = "", Optional ByVal pnMontMin As Currency = 0, Optional ByVal nSubDestino As Long = 0, Optional ByVal pnTpMonedad As Long = 0, Optional ByVal pnTpInteres As Long = 0)
cTpProd = pcTpProd
nMontMin = pnMontMin
nTpMonedad = pnTpMonedad
If pcTpProd = "803" And (nDestino = 27 Or nDestino = 28) Then
    lblDesc.Caption = "Monto"
    Me.Caption = "Monto de la Vivienda"
ElseIf pcTpProd = "703" Then
    
    lblMontoDisp.Visible = True
    editMMontoDisp.Visible = True
    
    If pnTpInteres = 48002 Then
       lblDesc.Caption = "Monto - Int. adelantado"
        Me.Caption = "Depósito de Plazo Fijo"
    Else
        lblDesc.Caption = "Monto"
        Me.Caption = "Depósito de Plazo Fijo"
    End If
Else
    If pcTpProd = 521 And nSubDestino = 35001 Then
        lblDesc.Caption = "Monto"
        Me.Caption = "Activo ó Terreno"
    ElseIf pcTpProd = 718 And nDestino = 15 Then
        lblDesc.Caption = "Monto"
        Me.Caption = "Terreno"
    Else
        lblDesc.Caption = "Monto"
        Me.Caption = "Presupuesto"
    End If
End If

If pcTpProd = "703" Then
    lblMontoDisp.Visible = True
    editMMontoDisp.Visible = True
    editMMontoSoli.Top = 1440
    lblMontSol.Top = 1440
    cmdGuardar.Top = 1800
    frDatos.Height = 2175
    stDatos.Height = 2775
    frmCredMontoPresupuestado.Height = 3255
Else
    lblMontoDisp.Visible = False
    editMMontoDisp.Visible = False
    editMMontoSoli.Top = 1200
    lblMontSol.Top = 1200
    cmdGuardar.Top = 1560
    frDatos.Height = 1935
    stDatos.Height = 2535
    frmCredMontoPresupuestado.Height = 3030
End If

nPorcentaje = pnPorcentaje
EditMPorcentaje = pnPorcentaje

If Not IsArray(pnMatMontoSol) Then
    Set nMatMontoSol = Nothing
Else
   nMatMontoSol = pnMatMontoSol
End If

Call CargaDatos
EnfocaControl editMMontoPresu
If pcTpProd = "703" Then
    MsgBox "Si corresponde, considerar saldo de créditos vigentes.", vbInformation, "Aviso"
End If
Me.Show 1
If IsArray(nMatMontoSol) Then
    pnMatMontoSol = nMatMontoSol
End If
End Sub

Private Sub Cmdguardar_Click()
If Valida() Then
    If cTpProd = "703" Then
        ReDim nMatMontoSol(1, 4)
        nMatMontoSol(1, 1) = editMMontoPresu
        nMatMontoSol(1, 2) = EditMPorcentaje
        nMatMontoSol(1, 3) = editMMontoSoli
        nMatMontoSol(1, 4) = Format((editMMontoPresu * (EditMPorcentaje / 100)), "#,#00.00")
    Else
        ReDim nMatMontoSol(1, 3)
        nMatMontoSol(1, 1) = editMMontoPresu
        nMatMontoSol(1, 2) = EditMPorcentaje
        nMatMontoSol(1, 3) = editMMontoPresu - editMMontoPresu * (EditMPorcentaje / 100)
    End If
    Unload Me
End If
End Sub
Private Sub CargaDatos()
If IsArray(nMatMontoSol) Then
    If UBound(nMatMontoSol) > 0 Then
        If cTpProd = "703" Then
            editMMontoPresu = Format(nMatMontoSol(1, 1), "#,#00.00")
            EditMPorcentaje = Format(nMatMontoSol(1, 2), "#0")
            editMMontoDisp = Format(nMatMontoSol(1, 4), "#,#00.00")
            editMMontoSoli = Format(nMatMontoSol(1, 3), "#,#00.00")
        Else
            editMMontoPresu = Format(nMatMontoSol(1, 1), "#,#00.00")
            EditMPorcentaje = Format(nMatMontoSol(1, 2), "#0")
            editMMontoSoli = Format((editMMontoPresu - (Format(editMMontoPresu, "#,#00.00") * (Format(EditMPorcentaje, "#0") / 100))), "#,#00.00")
        End If
    Else
        editMMontoPresu = nMatMontoSol(1, 1)
        If cTpProd = "703" Then
            editMMontoDisp = Format((editMMontoPresu * (EditMPorcentaje / 100)), "#,#00.00")
            editMMontoSoli = 0
        Else
            editMMontoSoli = Format((editMMontoPresu - (editMMontoPresu * (EditMPorcentaje / 100))), "#,#00.00")
        End If
    End If
End If
End Sub

Private Sub editMMontoPresu_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cTpProd = "703" Then
            editMMontoDisp = Format((editMMontoPresu * (EditMPorcentaje / 100)), "#,#00.00")
            editMMontoSoli = Format(editMMontoDisp, "#,#00.00")
        Else
            editMMontoSoli = Format((editMMontoPresu - (editMMontoPresu * (EditMPorcentaje / 100))), "#,#00.00")
        End If
        cmdGuardar.SetFocus
    End If
End Sub

Private Function Valida() As Boolean
Valida = True
    If editMMontoPresu = 0 Or editMMontoPresu = "" Then
        MsgBox "Ingrese el Monto", vbInformation, "Aviso"
        Valida = False
        editMMontoPresu.SetFocus
        Exit Function
    End If
    If CCur(editMMontoSoli) < CCur(nMontMin) Then
        MsgBox "El Monto mínimo es " & IIf(nTpMonedad = 1, "S/ ", "$ ") & Format(nMontMin, "#,#00.00"), vbInformation, "Aviso"
        Valida = False
        editMMontoPresu.SetFocus
        Exit Function
    End If
    If CCur(EditMPorcentaje) < nPorcentaje And cTpProd <> "703" Then
        MsgBox "El Aporte mínimo es " & Format(nPorcentaje, "#0") & "%", vbInformation, "Aviso"
        Valida = False
        EditMPorcentaje = Format(nPorcentaje, "#0")
        EditMPorcentaje.SetFocus
        Exit Function
    End If
    If CCur(EditMPorcentaje) <> nPorcentaje And cTpProd = "703" Then
        MsgBox "El porcentaje para este tipo de Interés es " & Format(nPorcentaje, "#0") & " %", vbInformation, "Aviso"
        Valida = False
        EditMPorcentaje.SetFocus
        Exit Function
    End If
    If CCur(editMMontoSoli) > CCur(editMMontoDisp) And cTpProd = "703" Then
        MsgBox "El monto solicitado no puede exceder al monto disponible", vbInformation, "Aviso"
        Valida = False
        editMMontoSoli.SetFocus
        Exit Function
    End If
End Function

Private Sub editMMontoPresu_LostFocus()
    If cTpProd = "703" Then
        editMMontoDisp = Format((editMMontoPresu * (EditMPorcentaje / 100)), "#,#00.00")
        editMMontoSoli = Format(editMMontoDisp, "#,#00.00")
    Else
        editMMontoSoli = Format((editMMontoPresu - (editMMontoPresu * (EditMPorcentaje / 100))), "#,#00.00")
    End If
End Sub

Private Sub editMMontoSoli_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdGuardar.SetFocus
    End If
End Sub

Private Sub EditMPorcentaje_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call editMMontoPresu_KeyPress(13)
    End If
End Sub

Private Sub EditMPorcentaje_LostFocus()
    Call editMMontoPresu_KeyPress(13)
End Sub

Private Sub Form_Load()
CentraForm Me
DisableCloseButton Me
End Sub

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
