VERSION 5.00
Begin VB.Form frmCredMiViviendaAlertas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alertas Créditos ""MIVIVIENDA"""
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   Icon            =   "frmCredMiViviendaAlertas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
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
      Left            =   2280
      TabIndex        =   2
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
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
      Left            =   3840
      TabIndex        =   1
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   $"frmCredMiViviendaAlertas.frx":030A
      Height          =   795
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   7110
   End
   Begin VB.Label Label2 
      Caption         =   $"frmCredMiViviendaAlertas.frx":03BA
      Height          =   795
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   7110
   End
   Begin VB.Label Label1 
      Caption         =   $"frmCredMiViviendaAlertas.frx":04BC
      Height          =   795
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7110
   End
End
Attribute VB_Name = "frmCredMiViviendaAlertas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private fbPasar As Boolean
Private fsCtaCod As String
Private fnMonto As Double
Private fDCredito As COMDCredito.DCOMCredito

Private Sub CmdAceptar_Click()
If Not Verificar(fsCtaCod, fnMonto) Then
    fbPasar = False
Else
    fbPasar = True
End If

Unload Me
End Sub

Private Sub cmdImprimir_Click()
Dim oDoc As cPDF
Set oDoc = New cPDF

'Creación del Archivo
oDoc.Author = UCase(gsCodUser)
oDoc.Creator = "SICMACT - Negocio"
oDoc.Producer = "Caja Municipal de Ahorro y Crédito de Maynas S.A."
oDoc.Subject = "Administración  de Alertas para Créditos ''MIVIVIENDA''"
oDoc.Title = "Administración  de Alertas para Créditos ''MIVIVIENDA''"

If Not oDoc.PDFCreate(App.Path & "\Spooler\AlertasMIVIVIENDA" & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
    Exit Sub
End If

oDoc.Fonts.Add "F1", "Arial", TrueType, Normal, WinAnsiEncoding
oDoc.Fonts.Add "F2", "Arial", TrueType, Bold, WinAnsiEncoding

oDoc.LoadImageFromFile App.Path & "\Logo_2015.jpg", "Logo"

'Tamaño de hoja A4
oDoc.NewPage A4_Vertical

oDoc.WImage 65, 480, 50, 100, "Logo"
oDoc.WTextBox 80, 50, 15, 500, "CARTA DE ACEPTACIÓN (Créditos MIVIVIENDA)", "F2", 11, hCenter

oDoc.WTextBox 150, 40, 60, 520, Label1.Caption, "F1", 10, hjustify
oDoc.WTextBox 200, 40, 60, 520, Label2.Caption, "F1", 10, hjustify
oDoc.WTextBox 250, 40, 60, 520, Label3.Caption, "F1", 10, hjustify

oDoc.WTextBox 400, 20, 60, 500, ArmaFecha(gdFecSis), "F1", 9, hRight
oDoc.WTextBox 420, 20, 60, 400, "Firma:", "F1", 9, hRight
oDoc.WTextBox 420, 20, 60, 500, "___________________", "F1", 9, hRight

oDoc.PDFClose
oDoc.Show
End Sub


Public Function Inicio(ByVal psCtaCod As String, ByVal pnMonto As Double, Optional ByVal MatDatVivienda As Variant)
fsCtaCod = psCtaCod
fnMonto = pnMonto
If Not Verificar(fsCtaCod, fnMonto, MatDatVivienda) Then
    Me.Show 1
Else
    fbPasar = True
End If
Inicio = fbPasar
End Function
Public Sub DarBajaCanPagoAnticipado(ByVal psCtaCod As String)
Set fDCredito = New COMDCredito.DCOMCredito
Call fDCredito.EliminarMiViviendaAlertasPago(psCtaCod, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser))
End Sub

Private Function Verificar(ByVal psCtaCod As String, ByVal pnMonto As Double, _
                            Optional ByVal MatDatVivienda As Variant) As Boolean
Set fDCredito = New COMDCredito.DCOMCredito
Dim rs As ADODB.Recordset
Verificar = False

Set rs = fDCredito.ObtenerCredAlertaMIVIVIENDACred(psCtaCod, pnMonto)

If Not (rs.EOF And rs.BOF) Then
    If CInt(rs!nEstado) = 1 Then
        Verificar = True
    Else
        Verificar = False
    End If
Else
    Call fDCredito.RegistroMiViviendaAlertasPago(psCtaCod, pnMonto, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), MatDatVivienda)
    Verificar = False
End If
End Function

Private Sub Form_Unload(Cancel As Integer)
If Not Verificar(fsCtaCod, fnMonto) Then
    fbPasar = False
Else
    fbPasar = True
End If
End Sub
