VERSION 5.00
Begin VB.Form frmColPSelectBoveda 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4005
   ClientLeft      =   2925
   ClientTop       =   2235
   ClientWidth     =   4575
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Pic2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   225
      ScaleHeight     =   300
      ScaleWidth      =   4095
      TabIndex        =   4
      Top             =   120
      Width           =   4095
   End
   Begin VB.PictureBox Pic1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3420
      Left            =   225
      ScaleHeight     =   3420
      ScaleWidth      =   735
      TabIndex        =   3
      Top             =   405
      Width           =   735
   End
   Begin VB.CommandButton cmdAceptar 
      Cancel          =   -1  'True
      Caption         =   "&Aceptar"
      Height          =   405
      Left            =   2265
      TabIndex        =   2
      Top             =   3480
      Width           =   1035
   End
   Begin VB.Frame fraContenedor 
      Caption         =   "Haga un Click en la Agencia a escoger "
      Height          =   2895
      Index           =   3
      Left            =   1200
      TabIndex        =   0
      Top             =   480
      Width           =   3165
      Begin VB.ListBox List1 
         Height          =   2535
         ItemData        =   "frmColPSelectBoveda.frx":0000
         Left            =   210
         List            =   "frmColPSelectBoveda.frx":0002
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   255
         Width           =   2760
      End
   End
End
Attribute VB_Name = "frmColPSelectBoveda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
'* SELECCIONA BOVEDA DE PIGNORATICIO
'Archivo:  frmColPSelectBoveda.frm
'LAYG   :  15/10/2001.
'Resumen:  Nos permite ingresar o actualizar los precios del oro con
'          que van a ser procesados los listado y planillas para subasta
Option Explicit

Dim vNomFrm As Form
Dim x As Integer

Private Sub CmdAceptar_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
    'RotateText 90, Pic1, "Times New Roman", 18, 150, 2550, " BOVEDAS  "
    'RotateText 0, Pic2, "Times New Roman", 10, 740, 10, " A      S E L E C C I O N A R  "
    Dim loCargaAg As COMDColocPig.DCOMColPFunciones
    Dim lrAgenc As ADODB.Recordset
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    
    Set loCargaAg = New COMDColocPig.DCOMColPFunciones
        Set lrAgenc = loCargaAg.dObtieneAgencias(True)
    Set loCargaAg = Nothing
    If lrAgenc Is Nothing Then
        MsgBox " No se encuentran las Agencias ", vbInformation, " Aviso "
    Else
        Me.List1.Clear
        With lrAgenc
            Do While Not .EOF
                List1.AddItem !cAgeCod & " " & Trim(!cAgeDescripcion)
                If !cAgeCod = gsCodAge Then
                    List1.Selected(List1.ListCount - 1) = True
                End If
                .MoveNext
            Loop
        End With
    End If
End Sub

Public Sub Inicio(pNomFrm As Form)
    Set vNomFrm = pNomFrm
End Sub
