VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmOpeReportes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Registros de Efectivo por Usuario (Billetaje)"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   6045
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   660
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   5850
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
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
         Left            =   4200
         TabIndex        =   4
         Top             =   180
         Width           =   1380
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
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
         Left            =   2640
         TabIndex        =   3
         Top             =   180
         Width           =   1380
      End
   End
   Begin VB.CommandButton CmdSelecAge 
      Caption         =   "&Agencias"
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
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1185
   End
   Begin VB.Frame fraFecha 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Fechas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   630
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   5775
      Begin MSMask.MaskEdBox txtFechaF 
         Height          =   300
         Left            =   2475
         TabIndex        =   1
         Top             =   225
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   300
         Left            =   480
         TabIndex        =   0
         Top             =   225
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Al:"
         Height          =   195
         Left            =   2205
         TabIndex        =   7
         Top             =   285
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Del:"
         Height          =   195
         Left            =   90
         TabIndex        =   6
         Top             =   278
         Width           =   285
      End
   End
End
Attribute VB_Name = "frmOpeReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private fMatAgencias() As String

Private Sub cmdImprimir_Click()
Dim lnContAge As Integer, i As Integer
Dim lsCadAge As String

    ReDim fMatAgencias(0)
    lnContAge = 0
    For i = 0 To frmSelectAgencias.List1.ListCount - 1
        If frmSelectAgencias.List1.Selected(i) = True Then
            lnContAge = lnContAge + 1
            ReDim Preserve fMatAgencias(lnContAge)
            fMatAgencias(lnContAge - 1) = Mid(frmSelectAgencias.List1.List(i), 1, 2)
            If Len(Trim(lsCadAge)) = 0 Then
                lsCadAge = Mid(frmSelectAgencias.List1.List(i), 1, 2)
            Else
                lsCadAge = lsCadAge & "," & Mid(frmSelectAgencias.List1.List(i), 1, 2)
            End If
        End If
    Next i
    If lnContAge = 0 Then
        ReDim fMatAgencias(1)
        fMatAgencias(0) = gsCodAge
    End If
    
    Call MostrarReporteRegistrosEfectivoPorUsuario(txtFecha.Text, txtFechaF.Text, lsCadAge)
End Sub

'**DAOR 20080125
'**Muestra el reporte de registros de efectivo por usuario
Public Sub MostrarReporteRegistrosEfectivoPorUsuario(ByVal psFechaI As Date, ByVal psFechaF As Date, ByVal psListaAgencias As String)
Dim loNCOMcajero As COMNCajaGeneral.NCOMCajero
Dim R As ADODB.Recordset
Dim lsNombreArchivo As String
Dim lMatCabecera As Variant
Dim lsMensaje As String

    lsNombreArchivo = "RegistrosEfectivo"
    
    ReDim lMatCabecera(14, 2)

    lMatCabecera(0, 0) = "Agencia"
    lMatCabecera(1, 0) = "Fecha"
    lMatCabecera(2, 0) = "Usuario"
    lMatCabecera(3, 0) = "Usuario Reg"
    lMatCabecera(4, 0) = "Nombre de Usuario"
    lMatCabecera(5, 0) = "Nº Reg. Efec."
    lMatCabecera(6, 0) = "Ingresos MN"
    lMatCabecera(7, 0) = "Egresos MN"
    lMatCabecera(8, 0) = "Efectivo MN"
    lMatCabecera(9, 0) = "Sobr/Falt MN"
    lMatCabecera(10, 0) = "Ingresos ME"
    lMatCabecera(11, 0) = "Egresos ME"
    lMatCabecera(12, 0) = "Efectivo ME"
    lMatCabecera(13, 0) = "Sobr/Falt ME"
    
    Set loNCOMcajero = New COMNCajaGeneral.NCOMCajero
    Set R = loNCOMcajero.ObtenerReporteRegistrosEfectivoPorUsuario(Format(psFechaI, "yyyymmdd"), Format(psFechaF, "yyyymmdd"), psListaAgencias)
    Set loNCOMcajero = Nothing
          
    Call GeneraReporteEnArchivoExcel(gsNomCmac, gsNomAge, gsCodUser, gdFecSis, "Reporte de Registros de Efectivo por Usuario", "Del " & Format(psFechaI, "dd/mm/yyyy") & " Al " & Format(psFechaF, "dd/mm/yyyy"), lsNombreArchivo, lMatCabecera, R, 2, , , True)
    
End Sub



Private Sub cmdSalir_Click()
    Unload Me
End Sub


Private Sub CmdSelecAge_Click()
Dim i As Integer
Dim lnContAge As Integer

    frmSelectAgencias.Show 1
    ReDim fMatAgencias(0)
    lnContAge = 0
    For i = 0 To frmSelectAgencias.List1.ListCount - 1
        If frmSelectAgencias.List1.Selected(i) = True Then
            lnContAge = lnContAge + 1
            ReDim Preserve fMatAgencias(lnContAge)
            fMatAgencias(lnContAge - 1) = Mid(frmSelectAgencias.List1.List(i), 1, 2)
        End If
    Next i
End Sub
