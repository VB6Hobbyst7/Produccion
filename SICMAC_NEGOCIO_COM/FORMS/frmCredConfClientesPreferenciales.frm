VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredConfClientesPreferenciales 
   Caption         =   "Configuración de Clientes Preferenciales"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9135
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCredConfClientesPreferenciales.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   9135
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   7858
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Configuración General"
      TabPicture(0)   =   "frmCredConfClientesPreferenciales.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdGuardar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdCerrar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.CommandButton cmdCerrar 
         Cancel          =   -1  'True
         Caption         =   "Cerrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   7320
         TabIndex        =   18
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "Guardar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   5880
         TabIndex        =   17
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Caption         =   "Parámetros Configurables"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   3135
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   8295
         Begin VB.TextBox txtEndeMaxUniFami 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   5400
            TabIndex        =   11
            Top             =   1640
            Width           =   1450
         End
         Begin VB.TextBox txtMinCalSbsNorm 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   5400
            TabIndex        =   10
            Top             =   1220
            Width           =   1450
         End
         Begin VB.TextBox txtTieMaxTotCre 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   5400
            TabIndex        =   9
            Top             =   800
            Width           =   1450
         End
         Begin VB.TextBox txtPerContMini 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   5400
            TabIndex        =   8
            Top             =   380
            Width           =   1450
         End
         Begin VB.TextBox txtEdadMaxClie 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   5400
            TabIndex        =   7
            Top             =   2060
            Width           =   1450
         End
         Begin VB.Label Label14 
            Caption         =   "Años"
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
            Left            =   6960
            TabIndex        =   16
            Top             =   2150
            Width           =   1095
         End
         Begin VB.Label Label13 
            Caption         =   "Entidades"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   200
            Left            =   6960
            TabIndex        =   15
            Top             =   1750
            Width           =   1200
         End
         Begin VB.Label Label12 
            Caption         =   "%"
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
            Left            =   6960
            TabIndex        =   14
            Top             =   1350
            Width           =   855
         End
         Begin VB.Label Label9 
            Caption         =   "Días"
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
            Left            =   6960
            TabIndex        =   13
            Top             =   960
            Width           =   1200
         End
         Begin VB.Label Label8 
            Caption         =   "Meses (Últimos)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   200
            Left            =   6960
            TabIndex        =   12
            Top             =   500
            Width           =   1200
         End
         Begin VB.Label Label7 
            Caption         =   "Edad Máxima Cliente (inclusive):"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   405
            TabIndex        =   6
            Top             =   2100
            Width           =   3495
         End
         Begin VB.Label Label6 
            Caption         =   "Endeudamiento Max Unidad Familiar:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   405
            TabIndex        =   5
            Top             =   1695
            Width           =   4335
         End
         Begin VB.Label Label5 
            Caption         =   "Mínima Calificación SBS Normal (%):"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   405
            TabIndex        =   4
            Top             =   1290
            Width           =   3975
         End
         Begin VB.Label Label2 
            Caption         =   "Tiempo Máximo Total entre créditos:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   405
            TabIndex        =   3
            Top             =   885
            Width           =   3855
         End
         Begin VB.Label Label1 
            Caption         =   "Periodo de Continuidad Mínima:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   405
            TabIndex        =   2
            Top             =   480
            Width           =   3375
         End
      End
   End
End
Attribute VB_Name = "frmCredConfClientesPreferenciales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim R As ADODB.Recordset

Private Sub cmdCerrar_Click()
Unload Me
End Sub
Private Sub cmdGuardar_Click()
    Dim oParam As COMDCredito.DCOMParametro
    Set oParam = New COMDCredito.DCOMParametro
    If MsgBox("Desea Guardar los Datos", vbInformation + vbYesNo, "Configuración de Clientes Preferenciales") = vbYes Then
        Call oParam.ModificarParametro(102746, "", Me.txtPerContMini.Text, "")
        Call oParam.ModificarParametro(102747, "", Me.txtTieMaxTotCre.Text, "")
        Call oParam.ModificarParametro(102748, "", Me.txtMinCalSbsNorm.Text, "")
        Call oParam.ModificarParametro(102749, "", Me.txtEndeMaxUniFami.Text, "")
        Call oParam.ModificarParametro(102750, "", Me.txtEdadMaxClie.Text, "")
        Set oParam = Nothing
        MsgBox "Los Datos se Guardaron"
        Call CargarDatos
    End If
End Sub
Sub CargarDatos()
    Dim oParam As COMDCredito.DCOMParametro
    Set oParam = New COMDCredito.DCOMParametro

    Me.txtPerContMini.Text = oParam.RecuperaValorParametro(102746)
    Me.txtTieMaxTotCre.Text = oParam.RecuperaValorParametro(102747)
    Me.txtMinCalSbsNorm.Text = oParam.RecuperaValorParametro(102748)
    Me.txtEndeMaxUniFami.Text = oParam.RecuperaValorParametro(102749)
    Me.txtEdadMaxClie.Text = oParam.RecuperaValorParametro(102750)
End Sub
Private Sub Form_Load()
    CentraForm Me
    Call CargarDatos
End Sub
Public Sub Registrar()
    Me.Show 1
End Sub

Private Sub txtPerContMini_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        Me.txtTieMaxTotCre.SetFocus
    End If
End Sub
Private Sub txtTieMaxTotCre_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        Me.txtMinCalSbsNorm.SetFocus
    End If
End Sub
Private Sub txtMinCalSbsNorm_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        If Me.txtMinCalSbsNorm.Text >= 0 And Me.txtMinCalSbsNorm <= 100 Then
                Me.txtEndeMaxUniFami.SetFocus
            Else
                MsgBox "El Maximo valor es hasta 100", vbInformation, "Aviso"
                Me.txtMinCalSbsNorm.SetFocus
        End If
    End If
End Sub
Private Sub txtEndeMaxUniFami_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        Me.txtEdadMaxClie.SetFocus
    End If
End Sub
Private Sub txtEdadMaxClie_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        If Me.txtEdadMaxClie.Text >= 18 Then
                Me.cmdGuardar.SetFocus
        Else
                MsgBox "La Edad Minimia es 18", vbInformation, "Aviso"
                Me.txtEdadMaxClie.SetFocus
        End If
    End If
End Sub
