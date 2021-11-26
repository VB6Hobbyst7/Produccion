VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmVisualizacionDesembolsos 
   Caption         =   "Visualizacion Desembolsos"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   9600
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Calendario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   -60
      TabIndex        =   1
      Top             =   1770
      Width           =   9645
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "Calendario de Pagos"
         Height          =   435
         Left            =   240
         TabIndex        =   17
         Top             =   4710
         Width           =   1575
      End
      Begin VB.Frame Frame3 
         Caption         =   "Resumen del Desembolso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   210
         TabIndex        =   12
         Top             =   3840
         Width           =   9225
         Begin VB.Label LblInteres 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3690
            TabIndex        =   16
            Top             =   300
            Width           =   1725
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Interes:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2880
            TabIndex        =   15
            Top             =   330
            Width           =   645
         End
         Begin VB.Label LblCapital 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   990
            TabIndex        =   14
            Top             =   330
            Width           =   1725
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Capital:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   210
            TabIndex        =   13
            Top             =   360
            Width           =   675
         End
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   8310
         TabIndex        =   2
         Top             =   4680
         Width           =   1095
      End
      Begin MSComctlLib.ListView Lst 
         Height          =   3540
         Left            =   180
         TabIndex        =   3
         Top             =   270
         Width           =   9285
         _ExtentX        =   16378
         _ExtentY        =   6244
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ColHdrIcons     =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nro Cuota"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Tipo"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fecha de Pago"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Monto de Capital"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Monto de Interes"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Cuota a Pagar"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Desembolsos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1755
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   9555
      Begin MSComctlLib.ListView LstRelacion 
         Height          =   780
         Left            =   4530
         TabIndex        =   19
         Top             =   870
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   1376
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ColHdrIcons     =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nombre"
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Relacion"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "RELACION CON LOS DESEMBOLSOS"
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
         Left            =   4560
         TabIndex        =   18
         Top             =   630
         Width           =   3645
      End
      Begin VB.Label LblTipoCuota 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   1890
         TabIndex        =   11
         Top             =   1230
         Width           =   2385
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Cuota:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   450
         TabIndex        =   10
         Top             =   1260
         Width           =   1320
      End
      Begin VB.Label LblPlazo 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   1890
         TabIndex        =   9
         Top             =   900
         Width           =   1365
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Plazo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1215
         TabIndex        =   8
         Top             =   930
         Width           =   555
      End
      Begin VB.Label LblTasa 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   1890
         TabIndex        =   7
         Top             =   570
         Width           =   1365
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tasa de Interes:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   315
         TabIndex        =   6
         Top             =   570
         Width           =   1455
      End
      Begin VB.Label LblTitular 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   1890
         TabIndex        =   5
         Top             =   210
         Width           =   7095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Titular del Credito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   195
         TabIndex        =   4
         Top             =   300
         Width           =   1575
      End
   End
End
Attribute VB_Name = "FrmVisualizacionDesembolsos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sCtaCod As String

Private Sub CmdDesembolso_Click()
    Dim oVisualizacion As COMNCredito.NCOMVisualizacion  'DVisualizacion
    Dim sDesAge As String
    Dim sCadImp As String
    Dim oPrevio As clsPrevio
    
    Set oVisualizacion = New COMNCredito.NCOMVisualizacion  'DVisualizacion
    sDesAge = oVisualizacion.ObtenerNomAge(gsCodAge)
    sCadImp = oVisualizacion.ImprimeComprobanteDesembolso(sCtaCod, sDesAge, gdFecSis, gsCodUser, "CMAC ICA")
    
    Set oVisualizacion = Nothing
    Set oPrevio = New clsPrevio
    oPrevio.Show sCadImp, "COMPROBANTE DE DESEMBOLSO"
End Sub

Private Sub CmdImprimir_Click()
    Dim oVisualizacion As COMNCredito.NCOMVisualizacion   'DVisualizacion
    Dim oPrevio As clsPrevio
    Dim rs_Report As ADODB.Recordset
    Dim sCadImpresion As String
    
        If Lst.ListItems.Count > 0 Then
            'Contruyendo datos para el reporte
            Set rs_Report = New ADODB.Recordset
            With rs_Report.Fields
                .Append "nCuota", adInteger
                .Append "Tipo", adVarChar, 20
                .Append "dVenc", adDate
                .Append "Capital", adDouble
                .Append "Interes", adDouble
            End With
            
            rs_Report.Open
            For i = 1 To Lst.ListItems.Count
                rs_Report.AddNew
                rs_Report(0) = Lst.ListItems(i).Text
                rs_Report(1) = Lst.ListItems(i).SubItems(1)
                rs_Report(2) = Lst.ListItems(i).SubItems(2)
                rs_Report(3) = Lst.ListItems(i).SubItems(3)
                rs_Report(4) = Lst.ListItems(i).SubItems(4)
                rs_Report.Update
            Next i
            Set oVisualizacion = New COMNCredito.NCOMVisualizacion  ' DVisualizacion
            sCadImpresion = oVisualizacion.ImpreDocumento(lbltitular.Caption, lbltasa.Caption, lblPlazo.Caption, lbltipocuota.Caption, _
                                          sCtaCod, gdFecSis, gsCodUser, rs_Report)
         End If
         
    Set oPrevio = New clsPrevio
    oPrevio.Show sCadImpresion, "", False
    Set oPrevio = Nothing
End Sub


Private Sub cmdsalir_Click()
    Unload Me
End Sub

Public Sub Inicio(ByVal psCtaCod As String)
 CargarDatos psCtaCod
 sCtaCod = psCtaCod
 Me.Show vbModal
End Sub

Sub CargarDatos(ByVal psCtaCod As String)
'    Dim rs As ADODB.Recordset
    Dim oVisualizacion As COMNCredito.NCOMVisualizacion   'DVisualizacion
'    Dim oCalend As COMNCredito.NCOMCalendario
    Dim iTem As ListItem
    Dim nCapital As Double
    Dim nInteres As Double
    
    Dim sTitular As String
    Dim rsDatos As ADODB.Recordset
    Dim rsRelac As ADODB.Recordset
    Dim rsCalend As ADODB.Recordset
    
    Set oVisualizacion = New COMNCredito.NCOMVisualizacion  'DVisualizacion
    Call oVisualizacion.CargarDatosVisualizacion(psCtaCod, gdFecSis, sTitular, rsDatos, rsRelac, rsCalend)
    Set oVisualizacion = Nothing
    lbltitular.Caption = sTitular
    
    'Set rs = oVisualizacion.DatosGenerales(psCtaCod)
    If Not rsDatos.EOF And Not rsDatos.BOF Then
        lbltasa.Caption = rsDatos!nTasaInteres & "%"
        If rsDatos!nPlazo > 0 Then
            lblPlazo.Caption = "Cada " & rsDatos!nPlazo & " dias"
        Else
            lblPlazo.Caption = "Los " & rsDatos!nPeriodoFechaFija & " de cada mes"
        End If
        
    End If
    lbltipocuota.Caption = rsDatos!cConsDescripcion
    
    'Set rs = oVisualizacion.RelacionDesembolso(psCtaCod)
    
    Do Until rsRelac.EOF
       Set iTem = LstRelacion.ListItems.Add(, , rsRelac!cPersNombre)
       iTem.SubItems(1) = rsRelac!cRelacion
       rsRelac.MoveNext
    Loop
    'Set rs = Nothing
    Set oVisualizacion = Nothing
    
    'Cargando el Calendario
'    Set oVisualizacion = New COMNCredito.NCOMVisualizacion  'DVisualizacion
    'Set rs = oVisualizacion.VerCalendario(psCtaCod)
    
'    If oVisualizacion.ChekingEqualsDateMoney(psCtaCod, gdFecSis) = True Then
'        Set rs = oVisualizacion.VerCalendarioD(psCtaCod, gdFecSis)
'    Else
'        Set oCalend = New COMNCredito.NCOMCalendario
'        Set rs = oVisualizacion.VerCalendario(psCtaCod)
'        Set oCalend = Nothing
'    End If
    
    nCapital = 0
    nInteres = 0
    'rsCalend.MoveFirst
    
    Do Until rsCalend.EOF
        Set iTem = Lst.ListItems.Add(, , rsCalend!nCuota)
        iTem.SubItems(1) = rsCalend!Tipo
        iTem.SubItems(2) = Format(rsCalend!dvenc, "dd/MM/yyyy")
        iTem.SubItems(3) = Format(rsCalend!capital, "#0.00")
        iTem.SubItems(4) = Format(rsCalend!Interes, "#0.00")
        iTem.SubItems(5) = Format(rsCalend!capital + rsCalend!Interes, "#0.00")
        If rsCalend!Tipo = "Desembolso" Then
            nCapital = nCapital + rsCalend!capital
        End If
        nInteres = nInteres + rsCalend!Interes
        rsCalend.MoveNext
    Loop
    
    'Set rs = Nothing
    
    If Mid(psCtaCod, 9, 1) = "1" Then
        lblCapital.Caption = Format(nCapital, "#0.00")
        lblInteres.Caption = Format(nInteres, "#0.00")
        lblCapital.ForeColor = vbBlue
        lblInteres.ForeColor = vbBlue
    Else
        lblCapital.Caption = Format(nCapital, "#0.00")
        lblInteres.Caption = Format(nInteres, "#0.00")
        lblCapital.ForeColor = vbGreen
        lblInteres.ForeColor = vbGreen
    End If
End Sub
