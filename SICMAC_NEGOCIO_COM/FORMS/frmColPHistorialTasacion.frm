VERSION 5.00
Begin VB.Form frmColPHistorialTasacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Historial "
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3930
   Icon            =   "frmColPHistorialTasacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   3930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   2760
      TabIndex        =   10
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label lblTasador 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2880
      TabIndex        =   12
      Top             =   1440
      Width           =   900
   End
   Begin VB.Label Label6 
      Caption         =   "TASADOR"
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
      Left            =   2880
      TabIndex        =   11
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblPesoNeto 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1440
      TabIndex        =   9
      Top             =   2280
      Width           =   795
   End
   Begin VB.Label lblPesoBruto 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1440
      TabIndex        =   8
      Top             =   1920
      Width           =   795
   End
   Begin VB.Label lblKilataje 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1440
      TabIndex        =   7
      Top             =   1560
      Width           =   795
   End
   Begin VB.Label lblPieza 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1440
      TabIndex        =   6
      Top             =   1200
      Width           =   795
   End
   Begin VB.Label lblDesc 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   3735
   End
   Begin VB.Label Label5 
      Caption         =   "PESO NETO:"
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
      TabIndex        =   4
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "PESO BRUTO:"
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
      TabIndex        =   3
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "KILATAJE:"
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
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "PIEZAS:"
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
      TabIndex        =   1
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "DESCRIPCION:"
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
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmColPHistorialTasacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre      : frmColPHistorialTasacion
'** Descripción : Formulario que muestra la tasacion de un credito prendario
'** Creación    : RECO, 20140707 - ERS074-2014
'**********************************************************************************************
Option Explicit
Private Sub cmdSalir_Click()
    Unload Me
End Sub

Public Sub Inicio(ByVal psDesc As String, ByVal pnPieza As Integer, ByVal pskilataje As String, ByVal pnPesoBruto As Double, ByVal pnPesoNeto As Double, ByVal psTasador As String)
    lblDesc.Caption = psDesc
    lblPieza.Caption = pnPieza
    lblKilataje.Caption = pskilataje
    lblPesoBruto.Caption = Format(pnPesoBruto, gcFormView)
    lblPesoNeto.Caption = Format(pnPesoNeto, gcFormView)
    lblTasador.Caption = psTasador
    Me.Show 1
End Sub
