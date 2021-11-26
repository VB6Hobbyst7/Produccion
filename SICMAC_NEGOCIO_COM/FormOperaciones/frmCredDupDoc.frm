VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCredDupDoc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Duplicados de Documentos"
   ClientHeight    =   7575
   ClientLeft      =   2670
   ClientTop       =   1875
   ClientWidth     =   7125
   Icon            =   "frmCredDupDoc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   75
      TabIndex        =   9
      Top             =   6810
      Width           =   6990
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   5640
         TabIndex        =   5
         Top             =   255
         Width           =   1260
      End
      Begin VB.CommandButton CmdNewBusq 
         Caption         =   "&Nueva Busqueda"
         Height          =   375
         Left            =   1515
         TabIndex        =   4
         Top             =   255
         Width           =   1710
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "&Imprimir"
         Enabled         =   0   'False
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
         Left            =   105
         TabIndex        =   3
         ToolTipText     =   "Imprima el Reporte Seleccionado"
         Top             =   270
         Width           =   1320
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4125
      Left            =   75
      TabIndex        =   8
      Top             =   2670
      Width           =   6990
      Begin MSComctlLib.ListView LstReportes 
         Height          =   3735
         Left            =   90
         TabIndex        =   2
         ToolTipText     =   "Seleccione un Reporte para su impresion"
         Top             =   255
         Width           =   6765
         _ExtentX        =   11933
         _ExtentY        =   6588
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483647
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Reportes"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Codigo"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Credito"
      Height          =   2610
      Left            =   75
      TabIndex        =   6
      Top             =   60
      Width           =   6990
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Buscar"
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
         TabIndex        =   0
         Top             =   315
         Width           =   900
      End
      Begin VB.Frame FraListaCred 
         Caption         =   "&Lista Creditos"
         Height          =   960
         Left            =   4800
         TabIndex        =   7
         Top             =   150
         Width           =   2115
         Begin VB.ListBox LstCred 
            Height          =   645
            ItemData        =   "frmCredDupDoc.frx":030A
            Left            =   75
            List            =   "frmCredDupDoc.frx":030C
            TabIndex        =   1
            Top             =   225
            Width           =   1980
         End
      End
      Begin SICMACT.ActXCodCta ActxCta 
         Height          =   420
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   3675
         _extentx        =   6482
         _extenty        =   741
         texto           =   "Credito :"
         enabledcmac     =   -1
         enabledcta      =   -1
         enabledprod     =   -1
         enabledage      =   -1
      End
      Begin VB.Label lblcodigo 
         BackColor       =   &H00FFFFFF&
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
         Height          =   300
         Index           =   0
         Left            =   1425
         TabIndex        =   20
         Top             =   1365
         Width           =   1200
      End
      Begin VB.Label lblcodigo 
         BackColor       =   &H00FFFFFF&
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
         Height          =   300
         Index           =   1
         Left            =   1425
         TabIndex        =   19
         Top             =   1695
         Width           =   4755
      End
      Begin VB.Label lblcodigo 
         BackColor       =   &H00FFFFFF&
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
         Height          =   300
         Index           =   2
         Left            =   1425
         TabIndex        =   18
         Top             =   2025
         Width           =   1230
      End
      Begin VB.Label lblcodigo 
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         Index           =   3
         Left            =   4695
         TabIndex        =   17
         Top             =   2055
         Width           =   1215
      End
      Begin VB.Label lblNat 
         AutoSize        =   -1  'True
         Caption         =   "Doc. Natural :"
         Height          =   195
         Left            =   270
         TabIndex        =   16
         Top             =   2070
         Width           =   990
      End
      Begin VB.Label lblTrib 
         AutoSize        =   -1  'True
         Caption         =   "Doc. Juridico :"
         Height          =   195
         Left            =   3270
         TabIndex        =   15
         Top             =   2100
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Left            =   255
         TabIndex        =   14
         Top             =   1410
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nombre :"
         Height          =   195
         Left            =   270
         TabIndex        =   13
         Top             =   1740
         Width           =   645
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         Caption         =   "Datos del Titular"
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
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   1065
         Width           =   1425
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Estado Crédito :"
         Height          =   195
         Left            =   2745
         TabIndex        =   11
         Top             =   1410
         Width           =   1125
      End
      Begin VB.Label lblcodigo 
         BackColor       =   &H00FFFFFF&
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
         Height          =   300
         Index           =   4
         Left            =   3915
         TabIndex        =   10
         Top             =   1350
         Width           =   1740
      End
      Begin VB.Line Line1 
         X1              =   105
         X2              =   6645
         Y1              =   1275
         Y2              =   1275
      End
   End
End
Attribute VB_Name = "frmCredDupDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nMiVivienda As Integer
Dim nPrestamo As Currency
Dim nCalendDinamico As Integer
Dim bCuotaCom As Integer
Dim lcDNI As String, lcRUC As String
Dim lnEstCred As Long '*** PEAC 20091211
Dim fbPersNatural As Boolean
Dim lsPersTDoc As Integer
'MADM 20110224
Dim oPersona As COMDPersona.UCOMPersona
Dim bvalorNegativo As Integer
Dim lnCondicion As Integer
'END MADM
'FRHU 20121202
'Dim oCliPre As COMNCredito.NCOMCredito 'COMENTADO POR ARLO 20170722
Dim bValidarCliPre As Boolean
'FIN FRHU 20121202
Dim bPermisoCargo As Boolean 'FRHU 20140722 ERS105-2014

Dim nValorR1 As Integer 'JOEP20170622
Dim nValorR2 As Integer 'JOEP20170622
Dim objProducto As COMDCredito.DCOMCredito '**ARLO20180712 ERS042 - 2018

Dim MatPersona(1) As TActAutDatos 'add pti1 ers070-2018
Private Enum TiposBusquedaNombre
    BusqApellidoPaterno = 1
    BusqApellidoMaterno = 2
    BusqApellidoCasada = 3
    BusqNombres = 4
End Enum
Dim ultEstado As Integer
Dim sAgeReg As String
Dim dfreg As String
'fin add pti1 ers070-2018




'RECO20150331 ERS010-2015****************************
Public Sub Iniciaformulario(ByVal psCtaCod As String)
    Me.ActxCta.NroCuenta = psCtaCod
    Call ActxCta_KeyPress(vbKeyReturn)
    Me.Show 1
    
End Sub
'RECO FIN *******************************************

Private Function CargaDatos(ByVal psCtaCod As String) As Boolean
'Dim sSQL As String
'Dim R As ADODB.Recordset
Dim oCred As COMDCredito.DCOMCredito
'Dim nEstado As Integer
Dim rsCred As ADODB.Recordset
Dim rsComun As ADODB.Recordset
Dim bCargado As Boolean
Dim sEstado As String

    Set oCred = New COMDCredito.DCOMCredito
    bCargado = oCred.CargaDatosReportesDuplicados(psCtaCod, rsCred, rsComun, sEstado, lnEstCred)
    Set oCred = Nothing

    'Set R = oCred.RecuperaProducto(psCtaCod)
    'Set oCred = Nothing
    'If Not R.BOF And Not R.EOF Then
    If bCargado Then
    '    CargaDatos = True
        lblcodigo(4).Caption = sEstado 'ValorConstante(gColocEstado, R!nPrdEstado)
    '    nEstado = R!nPrdEstado
    '    R.Close
    '    Set R = Nothing
    '    Set R = oCred.RecuperaDatosComunes(psCtaCod, , Array(nEstado, 0))
        nPrestamo = IIf(IsNull(rsComun!nMontoCol), rsComun!nMontoSol, rsComun!nMontoCol)
        If nPrestamo = 0 Then
            nPrestamo = rsComun!nMontoSol
        End If
        nMiVivienda = IIf(IsNull(rsComun!bMiVivienda), 0, rsComun!bMiVivienda)
        bCuotaCom = IIf(IsNull(rsComun!bCuotaCom), 0, rsComun!bCuotaCom)
        nCalendDinamico = IIf(IsNull(rsComun!nCalendDinamico), 0, rsComun!nCalendDinamico)
        lblcodigo(0) = rsComun!cperscod
        lblcodigo(1) = PstaNombre(rsComun!cTitular)
        lblcodigo(2) = IIf(IsNull(rsComun!DNI), "", rsComun!DNI)
        lblcodigo(3) = IIf(IsNull(rsComun!Ruc), "", rsComun!Ruc)

        lcDNI = Trim(IIf(IsNull(rsComun!DNI), "", rsComun!DNI))
        lcRUC = Trim(IIf(IsNull(rsComun!Ruc), "", rsComun!Ruc))
        
        '**MADM 20100107 *****************************************
        If rsComun!nPersPersoneria = 1 Then
            fbPersNatural = True
            lsPersTDoc = 1
            If Trim(lcDNI) = "" Then
                lsPersTDoc = CInt((rsComun!nTipoId))
            End If
        Else
            fbPersNatural = False
            lsPersTDoc = 3
        End If
        '*********************************************************

    '    R.Close
    '    Set R = Nothing
    'Else
    '    CargaDatos = False
    End If

    CargaDatos = bCargado
    'Set oCred = Nothing

End Function

Private Sub ActxCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        lblcodigo(0).Caption = ""
        lblcodigo(4).Caption = ""
        lblcodigo(1).Caption = ""
        lblcodigo(2).Caption = ""
        lblcodigo(3).Caption = ""
        If CargaDatos(ActxCta.NroCuenta) Then
            CmdImprimir.Enabled = True
        Else
            CmdImprimir.Enabled = False
        End If
    End If
End Sub

Private Sub cmdBuscar_Click()
Dim oCredito As COMDCredito.DCOMCredito
Dim R As ADODB.Recordset
Dim oPers As COMDPersona.UCOMPersona
    
    LstCred.Clear
    Set oPers = frmBuscaPersona.Inicio()
    If Not oPers Is Nothing Then
        Set oCredito = New COMDCredito.DCOMCredito
        Set R = oCredito.RecuperaCreditosVigentes(oPers.sPersCod, , Array(2021, gColocEstVigMor, gColocEstVigNorm, gColocEstVigNorm, gColocEstRefMor, gColocEstRefNorm, 2001, 2002, 2000))
        Do While Not R.EOF
            LstCred.AddItem R!cCtaCod
            R.MoveNext
        Loop
        R.Close
        Set R = Nothing
        Set oCredito = Nothing
    End If
    If LstCred.ListCount = 0 Then
        MsgBox "El Cliente No Tiene Creditos Vigentes", vbInformation, "Aviso"
    End If
    
End Sub


Public Sub ImprimePagareCred(ByVal psCtaCod As String, ByVal pnFormato As Integer)
Dim ssql As String
Dim oDCred As COMDCredito.DCOMCredito
Dim R As ADODB.Recordset
Dim RRelaCred As ADODB.Recordset
Dim sCadImp As String
Dim oFun As COMFunciones.FCOMCadenas
Dim nGaran As Integer
Dim nTitu As Integer
Dim nCode As Integer

Dim rsUbi As ADODB.Recordset 'MAVM 26112009 ***

Dim oWord As Word.Application
Dim oDoc As Word.Document
Dim oRange As Word.Range

Dim ObjCons As New COMDConstantes.DCOMAgencias
Dim sNomAgencia As String
    sNomAgencia = ObjCons.NombreAgencia(Mid(psCtaCod, 4, 2))
Set ObjCons = Nothing

Dim nTasaCompAnual As Double
Set oDCred = New COMDCredito.DCOMCredito
Set RRelaCred = oDCred.RecuperaRelacPers(psCtaCod)
    
    Dim ObjGarantes As New COMDCredito.DCOMCredActBD
    Dim sEmision As String
    Dim RsGarantes As New ADODB.Recordset
    
    Dim sArchivo As String
    
    'sEmision = oDCred.RecuperaNomUbigeoAgencia(psCtaCod)
    Set RsGarantes = oDCred.RecuperaGarantes(psCtaCod)
    
    Set oDCred = New COMDCredito.DCOMCredito
    Set R = oDCred.RecuperaDatosComunes(psCtaCod)
    
    'MAVM 26112009 *********
    Set rsUbi = oDCred.RecuperaUbigeo(Mid(psCtaCod, 4, 2))
    sEmision = rsUbi!cUbiGeoDescripcion
    
    Dim liPosicion As Integer
    Dim lsCiudad As String
    lsCiudad = Trim(sEmision)
    liPosicion = InStr(lsCiudad, "(")
    If liPosicion > 0 Then
    lsCiudad = Left(lsCiudad, liPosicion - 1)
    End If
    '*********
    
    Set oDCred = Nothing
    Set oFun = New COMFunciones.FCOMCadenas
  
    Set oWord = CreateObject("Word.Application")
        oWord.Visible = False
    If pnFormato = 0 Then
        Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\PAGARE.doc")
    Else
        Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\PAGARE2.doc")
    End If
    Set oDCred = Nothing
    
    sArchivo = App.Path & "\FormatoCarta\PAGARE_" & psCtaCod & ".doc"
    oDoc.SaveAs (sArchivo)
    
    'Set RsTasa = ObjTasa.RecuperaProductoTasaInteres(psCtaCod, gColocLineaCredTasasIntCompNormal)
    With oWord.Selection.Find
        .Text = "<<CREDITO>>"
        .Replacement.Text = psCtaCod
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    

    'ARCV 30-01-2007

'    Select Case Mid(psCtaCod, 4, 2)
'        Case "01", "09", "04"
'            sEmision = "IQUITOS"
'        Case "02"
'            sEmision = "HUANUCO"
'        Case "06"
'            sEmision = "YURIMAGUAS"
'        Case "07"
'            sEmision = "TINGO MARIA"
'        Case "06"
'            sEmision = "PUCALLPA"
'    End Select
    '--------
    
    With oWord.Selection.Find
        .Text = "<<LUGAR>>"
        .Replacement.Text = lsCiudad
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    With oWord.Selection.Find
        .Text = "<<OFICINA>>"
        .Replacement.Text = sNomAgencia
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    
    If Not (R.EOF And R.BOF) Then
        With oWord.Selection.Find
            .Text = "<<FECVENC>>"
            .Replacement.Text = Format(R!dVencPagare, "DD/MM/YYYY")
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        
        With oWord.Selection.Find
            .Text = "<<FECEMI>>"
            .Replacement.Text = Format(R!dVigencia, "DD/MM/YYYY")
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
                
        With oWord.Selection.Find
            .Text = "<<IMPORTE>>"
            .Replacement.Text = Format(R!nMontoPagare, "#0.00") 'R!nMontoCol 'ARCV 29-01-2007
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        
        With oWord.Selection.Find
            .Text = "<<MORATORIO>>"
            .Replacement.Text = R!nTasaMora
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        
        nTasaCompAnual = Format(((1 + R!nTasaInteres / 100) ^ (360 / 30) - 1) * 100, "#.00")
        
        With oWord.Selection.Find
            .Text = "<<TASAANUAL>>"
            .Replacement.Text = nTasaCompAnual
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        
       With oWord.Selection.Find
          .Text = "<<LETRAS>>"
          
          '.Replacement.Text = UnNumero(R!nMontoPagare) & " y 00/100" 'UnNumero(R!nMontoCol) 'ARCV 29-01-2007
          '.Replacement.Text = NumLet(R!nMontoPagare) & IIf(Mid(psCtaCod, 9, 1) = "2", "", " y " & IIf(InStr(1, R!nMontoPagare, ".") = 0, "00", Mid(R!nMontoPagare, InStr(1, R!nMontoPagare, ".") + 1, 2)) & "/100")
          .Replacement.Text = NumLet(R!nMontoPagare) & IIf(Mid(psCtaCod, 9, 1) = "2", "", " y " & IIf(InStr(1, R!nMontoPagare, ".") = 0, "00", Left(Mid(R!nMontoPagare, InStr(1, R!nMontoPagare, ".") + 1, 2) + "00", 2)) & "/100") 'EJVG20130924
          .Forward = True
          .Wrap = wdFindContinue
          .Format = False
          .Execute Replace:=wdReplaceAll
       End With
     End If
    

     
     With oWord.Selection.Find
        .Text = "<<MONEDA>>"
        .Replacement.Text = IIf(Mid(psCtaCod, 9, 1) = 1, "NUEVOS SOLES", "DOLARES")
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
     End With
     
    If Not (RRelaCred.EOF And RRelaCred.BOF) Then
       Do Until RRelaCred.EOF
          If RRelaCred!nConsValor = gColRelPersTitular Then
               With oWord.Selection.Find
                    .Text = "<<TITULAR>>"
                    .Replacement.Text = PstaNombre(RRelaCred!cPersNombre)
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
               End With
                
               With oWord.Selection.Find
                    .Text = "<<DOCTITULAR>>"
                    .Replacement.Text = IIf(IsNull(RRelaCred!DNI), RRelaCred!Ruc, RRelaCred!DNI)
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
               End With
                              
               With oWord.Selection.Find
                    .Text = "<<DIRECTITULAR>>"
                    .Replacement.Text = RRelaCred!cPersDireccDomicilio
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
               End With
               nTitu = 1
          'ARCV 07-05-2007
          'ElseIf RRelaCred!nConsValor = gColRelPersConyugue Then
          ElseIf RRelaCred!nConsValor = gColRelPersConyugue Or RRelaCred!nConsValor = gColRelPersCodeudor Then
          '-------------
               With oWord.Selection.Find
                    .Text = "<<CONYUGE>>"
                    .Replacement.Text = PstaNombre(RRelaCred!cPersNombre)
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
               End With
                
               With oWord.Selection.Find
                    .Text = "<<DOCCONYUGE>>"
                    .Replacement.Text = IIf(IsNull(RRelaCred!DNI), RRelaCred!Ruc, RRelaCred!DNI)
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
               End With
               
               With oWord.Selection.Find
                    .Text = "<<DIRECCONYUGE>>"
                    .Replacement.Text = RRelaCred!cPersDireccDomicilio
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
               End With
               nCode = 1
           Exit Do  'ARCV 07-05-2007
           
           End If
           
          'End If
          RRelaCred.MoveNext
       Loop
    End If
    
    With oWord.Selection.Find
         .Text = "<<CONYUGE>>"
         .Replacement.Text = ""
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
    End With
     
    With oWord.Selection.Find
         .Text = "<<DOCCONYUGE>>"
         .Replacement.Text = ""
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
    End With
    
    With oWord.Selection.Find
         .Text = "<<DIRECCONYUGE>>"
         .Replacement.Text = ""
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
    End With
    
    
    nGaran = 0
    If Not (RsGarantes.EOF And RsGarantes.BOF) Then
        While Not RsGarantes.EOF
            If nGaran = 0 Then
               With oWord.Selection.Find
                    .Text = "<<AVAL1>>"
                    .Replacement.Text = PstaNombre(RsGarantes!cPersNombre)
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
               End With
            
               With oWord.Selection.Find
                    .Text = "<<DIRECAVAL1>>"
                    .Replacement.Text = RsGarantes!cPersDireccDomicilio
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
               End With
               
               With oWord.Selection.Find
                    .Text = "<<DOCAVAL1>>"
                    .Replacement.Text = IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc)
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
               End With
               
            ElseIf nGaran = 1 Then
            
               With oWord.Selection.Find
                    .Text = "<<AVAL2>>"
                    '.Replacement.Text = RsGarantes!cPersNombre
                    .Replacement.Text = PstaNombre(RsGarantes!cPersNombre) 'EJVG20130122
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
               End With
            
               With oWord.Selection.Find
                    .Text = "<<DIRECAVAL2>>"
                    .Replacement.Text = RsGarantes!cPersDireccDomicilio
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
               End With
               
               With oWord.Selection.Find
                    .Text = "<<DOCAVAL2>>"
                    .Replacement.Text = IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc)
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
               End With
            
            End If
            nGaran = nGaran + 1
            RsGarantes.MoveNext
        Wend
    Else
    End If
    
    With oWord.Selection.Find
        .Text = "<<AVAL1>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
   End With

   With oWord.Selection.Find
        .Text = "<<DIRECAVAL1>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
   End With
   
   With oWord.Selection.Find
        .Text = "<<DOCAVAL1>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
   End With
   'AVAL2
   With oWord.Selection.Find
        .Text = "<<AVAL2>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
   End With

   With oWord.Selection.Find
        .Text = "<<DIRECAVAL2>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
   End With
   
   With oWord.Selection.Find
        .Text = "<<DOCAVAL2>>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
   End With

    
    Set RsGarantes = Nothing
    Set ObjGarantes = Nothing
    
    'ARCV 07-05-2007
    'oDoc.SaveAs (App.path & "\FormatoCarta\PAGARE_" & psCtaCod & ".doc")
    '--------
    oDoc.Close
    Set oDoc = Nothing
    
    Set oWord = CreateObject("Word.Application")
        oWord.Visible = True
    'ARCV 07-05-2007
    'Set oDoc = oWord.Documents.Open(App.path & "\FormatoCarta\PAGARE_" & psCtaCod & ".doc")
    Set oDoc = oWord.Documents.Open(sArchivo)
    '----
    Set oDoc = Nothing
    Set oWord = Nothing
    
End Sub

Private Sub cmdImprimir_Click()
Dim oCredDoc As COMNCredito.NCOMCredDoc
Dim sCadImp As String
Dim sCadImp_2 As String
Dim Prev As previo.clsprevio
Dim nTipoFormato As Integer
Dim sFecDes As String '**************Agregado por PASI20131127 segun TI-ERS136-2013
Dim bValor As Boolean '**************Agregado por PASI20131128 segun TI-ERS136-2013

Dim oCred As COMNCredito.NCOMCredDoc
Dim oCredB As COMNCredito.NCOMCredDoc
Dim oCredC As COMDCredito.DCOMCredito
Dim oDCred As COMDCredito.DCOMCredito
Dim oDCredC As COMDCredito.DCOMCredito
Dim oDCredSbs As COMDCredito.DCOMCredito

Dim oDCredDoc As COMDCredito.DCOMCredito
Dim oDGarantia As COMDCredito.DCOMGarantia
    
Dim oDAgencia As COMDConstantes.DCOMAgencias
Dim oDCredito As COMDCredito.DCOMCredDoc

Dim rsDatosDDJJ As ADODB.Recordset
Dim rsGarantTitular As ADODB.Recordset
Dim rsGarantAvales As ADODB.Recordset
Dim rsAgencia As ADODB.Recordset

Dim RDatFin As ADODB.Recordset
Dim RDatFin1 As ADODB.Recordset
Dim R As ADODB.Recordset
Dim RB As ADODB.Recordset
Dim RSbs As ADODB.Recordset
Dim RCap As ADODB.Recordset

Dim RRel As ADODB.Recordset
Dim R2 As ADODB.Recordset
Dim RCaliSbs As ADODB.Recordset
Dim RRelaCred As ADODB.Recordset
Dim REstadosCred As ADODB.Recordset
Dim RGarantCred As ADODB.Recordset

Dim oScore As COMDCredito.DCOMCredito 'EAAS20181120 SEGUN ERS-072-2018
Set oScore = New COMDCredito.DCOMCredito 'EAAS20181120 SEGUN ERS-072-2018
Dim RProtestosSinAclarar As ADODB.Recordset 'EAAS20181120 SEGUN ERS-072-2018
Dim RCobranzaCoativaSunat As ADODB.Recordset 'EAAS20181120 SEGUN ERS-072-2018
Dim RObligacionesCerradas As ADODB.Recordset 'EAAS20181120 SEGUN ERS-072-2018
Dim rsInforme As ADODB.Recordset 'EAAS20181120 SEGUN ERS-072-2018
Dim rsRelFinales As ADODB.Recordset 'EAAS20181120 SEGUN ERS-072-2018
Dim cParamExtraXML As String 'EAAS20181120 SEGUN ERS-072-2018
Dim cParamExtraJSON As String 'EAAS20181120 SEGUN ERS-072-2018
Dim bPasarParamExperian As Boolean 'EAAS20181120 SEGUN ERS-072-2018
Dim iCR As Integer 'EAAS20181120 SEGUN ERS-072-2018
Dim nProtestosSinAclarar As Integer 'EAAS20181120 SEGUN ERS-072-2018
Dim nCobranzaCoativaSunat As Integer 'EAAS20181120 SEGUN ERS-072-2018
Dim nObligacionesCerradas As Integer 'EAAS20181120 SEGUN ERS-072-2018
Set rsRelFinales = oScore.ObtenerRelacionesFinalesExperian(ActxCta.NroCuenta) 'EAAS20181120 SEGUN ERS-072-2018
Dim cCodigosTitular As String 'EAAS20181120 SEGUN ERS-072-2018

Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval '*** LUCV20160812, Según ERS004-2016
Dim rsInfVisita As ADODB.Recordset                    '*** LUCV20160812, Según ERS004-2016
Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval '*** LUCV20160812, Según ERS004-2016

'*********CTI3 11092018
Dim oDCredtpP As COMDCredito.DCOMCredito
Dim tpoPersJur As ADODB.Recordset

Dim oDCredAdc As COMDCredito.DCOMCredito
Dim RDatAdc As ADODB.Recordset

'Accionistas
Dim oDCredAdcA As COMDCredito.DCOMCredito
Dim RDatAdcA As ADODB.Recordset

'Directorio
Dim oDCredAdcD As COMDCredito.DCOMCredito
Dim RDatAdcD As ADODB.Recordset

'Gerencia
Dim oDCredAdcG As COMDCredito.DCOMCredito
Dim RDatAdcG As ADODB.Recordset

'Patrimonio
Dim oDCredAdcP As COMDCredito.DCOMCredito
Dim RDatAdcP As ADODB.Recordset

'Cargos
Dim oDCredAdcCa As COMDCredito.DCOMCredito
Dim RDatAdcCa As ADODB.Recordset
'***********

    'On Error GoTo ErrorCmdImprimir_Click
    'RECO20150331************************************
    Screen.MousePointer = 11
    'RECO FIN****************************************
    Select Case LstReportes.SelectedItem.SubItems(1)
        Case "001" 'Registro de Solicitud de Credito
            Set oCredDoc = New COMNCredito.NCOMCredDoc
            sCadImp = oCredDoc.ImprimeRegistroSolicitudDuplicado(ActxCta.NroCuenta, gsNomAge, gdFecSis, gsCodUser, gsNomCmac)
            Set oCredDoc = Nothing
            Set Prev = New clsprevio
            Prev.Show sCadImp, "", False
            Set Prev = Nothing
            
        Case "002" 'Solicitud de Credito
        
            Set oCredDoc = New COMNCredito.NCOMCredDoc
            Set Prev = New clsprevio
            Prev.Show oCredDoc.ImprimeSolicitud(ActxCta.NroCuenta, gsNomAge, Format(gdFecSis, "dd/mm/yyyy"), gsCodUser), "", False
            Set Prev = Nothing
            Set oCredDoc = Nothing
            
            'Add pti1 ers070-2018 ******************************
            Dim spercod As String
            Dim snumcuenta As String
            Dim oNPersona As New COMNPersona.NCOMPersona
            Dim rs As ADODB.Recordset
            Dim rsPersona As ADODB.Recordset
            Dim CondicionCred  As String
            Dim cperscod  As String
            Dim CantImpr As Integer
            Dim CanR As Integer
            Dim Tsi As Integer
            Dim Tno As Integer
            
            
            Dim sUbicGeografica As String
            Dim sNombres As String, sApePat As String, sApeMat As String, sApeCas As String
            Dim sSexo As String, sEstadoCivil As String
            Dim sDomicilio As String, cNacionalidad As String, sRefDomicilio As String
            Dim sTelefonos As String, sCelular As String, sEmail As String
            Dim sPersIDTpo As String, sPersIDnro As String
            
            snumcuenta = Trim(ActxCta.NroCuenta)
            Set rs = oNPersona.DupDocCred(snumcuenta)
            If Not (rs.EOF And rs.BOF) Then
                 CondicionCred = rs!CondicionCred
                 cperscod = rs!cperscod
                 CantImpr = rs!CantImpr
                 CanR = rs!CanR
                 Tsi = rs!Tsi
                 Tno = rs!Tno
              
           
                 
                 
                 Set rsPersona = oNPersona.ObtenerDatosParaActAutDeCliente(cperscod)
                If Not (rsPersona.EOF And rsPersona.BOF) And CanR > 0 Then
                    ultEstado = rs!ultEstado
                    sAgeReg = rs!Agencia
                    dfreg = rs!freg
                  
                    sApePat = BuscaNombre(rsPersona!cPersNombre, BusqApellidoPaterno)
                    sApeMat = BuscaNombre(rsPersona!cPersNombre, BusqApellidoMaterno)
                    sNombres = BuscaNombre(rsPersona!cPersNombre, BusqNombres)
                    sApeCas = BuscaNombre(rsPersona!cPersNombre, BusqApellidoCasada)
                    sSexo = Trim(IIf(IsNull(rsPersona!cPersnatSexo), "", rsPersona!cPersnatSexo))
                    sEstadoCivil = Trim(IIf(IsNull(rsPersona!nPersNatEstCiv), "", rsPersona!nPersNatEstCiv))
                    sDomicilio = rsPersona!cPersDireccDomicilio
                    sUbicGeografica = rsPersona!cPersDireccUbiGeo
                    sTelefonos = IIf(IsNull(rsPersona!cPersTelefono), "", rsPersona!cPersTelefono)
                    sCelular = IIf(IsNull(rsPersona!cPersCelular), "", rsPersona!cPersCelular)
                    sEmail = IIf(IsNull(rsPersona!cEmail), "", rsPersona!cEmail)
                    cNacionalidad = Trim(IIf(IsNull(rsPersona!cNacionalidad), "", rsPersona!cNacionalidad))
                    sRefDomicilio = Trim(IIf(IsNull(rsPersona!cPersRefDomicilio), "", rsPersona!cPersRefDomicilio))
                    sPersIDTpo = Trim(IIf(IsNull(rsPersona!cPersIDTpo), "", rsPersona!cPersIDTpo))
                    sPersIDnro = Trim(IIf(IsNull(rsPersona!cPersIDnro), "", rsPersona!cPersIDnro))
                    
                       MatPersona(1).sNombres = sNombres
                       MatPersona(1).sApePat = sApePat
                       MatPersona(1).sApeMat = sApeMat
                       MatPersona(1).sApeCas = sApeCas
                       MatPersona(1).sPersIDTpo = sPersIDTpo
                       MatPersona(1).sPersIDnro = sPersIDnro
                       MatPersona(1).sSexo = sSexo
                       MatPersona(1).sEstadoCivil = sEstadoCivil
                       MatPersona(1).cNacionalidad = cNacionalidad
                       MatPersona(1).sDomicilio = sDomicilio
                       MatPersona(1).sRefDomicilio = sRefDomicilio
                       MatPersona(1).sUbicGeografica = sUbicGeografica
                       MatPersona(1).sCelular = sCelular
                       MatPersona(1).sTelefonos = sTelefonos
                       MatPersona(1).sEmail = sEmail
                     If CantImpr = 0 Then
                           
                            
                             If CondicionCred = 1 Then
                                'Cliente Nuevo
                                Call ImprimirPdfCartillaAutorizacion
                                Call oNPersona.imprimeCartAut(gsCodUser, snumcuenta)
                             Else
                                If CanR = 1 Then
                                    'si es un cliente recurrente y por primera vez autoriza sus datos
                                    Call ImprimirPdfCartillaAutorizacion
                                    Call oNPersona.imprimeCartAut(gsCodUser, snumcuenta)
                                    
                                Else
                                  'si es un cliente recurrente y CAMBIA DE NO A SI
                                  If CanR > 1 And Tsi = 1 And ultEstado = 1 Then
                                    Call ImprimirPdfCartillaAutorizacion
                                    Call oNPersona.imprimeCartAut(gsCodUser, snumcuenta)
                                  End If
                                
                                End If
                             End If
                     Else
                        
                         If vbYes = MsgBox("¿Desea Re-Imprimir la cartilla de Autorización de Uso de Datos?", vbInformation + vbYesNo) Then
                                    Call ImprimirPdfCartillaAutorizacion
                                    Call oNPersona.imprimeCartAut(gsCodUser, snumcuenta)
                         End If
                        
                    End If
                End If
                
           End If
            
            
            
            
        Case "003" 'Resumen de Comite
            Set oCredDoc = New COMNCredito.NCOMCredDoc
            Set Prev = New clsprevio
            Prev.Show oCredDoc.ImprimeResumenComite(ActxCta.NroCuenta, gsNomAge, Format(gdFecSis, "dd/mm/yyyy"), gsCodUser, gsNomCmac), "", False
            Set Prev = Nothing
            Set oCredDoc = Nothing
            
        Case "004" 'Comprobante de Desembolso
            Set oCredDoc = New COMNCredito.NCOMCredDoc
            Set Prev = New clsprevio
            Prev.Show oCredDoc.ImprimeComprobanteDesembolso(ActxCta.NroCuenta, gsNomAge, Format(gdFecSis, "dd/mm/yyyy"), gsCodUser, gsNomCmac), "", False
            Set Prev = Nothing
            Set oCredDoc = Nothing
            
        Case "005" 'Plan de Pagos
            Set oCredDoc = New COMNCredito.NCOMCredDoc
            Set Prev = New clsprevio
            Dim bDetallado As Boolean
            bDetallado = IIf(MsgBox("Desea imprimir el Plan de Pagos Detallado?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmación") = vbYes, True, False)
            'Modificado por PASI20131128 segun TI-ERS136-2013
            'sCadImp = oCredDoc.ImprimePlandePagosDuplicado(ActxCta.NroCuenta, gsNomAge, Format(gdFecSis, "dd/mm/yyyy"), gsCodUser, nPrestamo, nMiVivienda, , gsNomCmac, bCuotaCom, nCalendDinamico, nMiVivienda, bDetallado)
            bValor = VerificarExisteDesembolsoBcoNac(ActxCta.NroCuenta, sFecDes, 1)
            If bValor = True Then
                sCadImp = oCredDoc.ImprimePlandePagosDuplicado(ActxCta.NroCuenta, gsNomAge, Format(gdFecSis, "dd/mm/yyyy"), gsCodUser, nPrestamo, nMiVivienda, , gsNomCmac, bCuotaCom, nCalendDinamico, nMiVivienda, bDetallado, sFecDes)
            Else
                sCadImp = oCredDoc.ImprimePlandePagosDuplicado(ActxCta.NroCuenta, gsNomAge, Format(gdFecSis, "dd/mm/yyyy"), gsCodUser, nPrestamo, nMiVivienda, , gsNomCmac, bCuotaCom, nCalendDinamico, nMiVivienda, bDetallado)
            End If
            'End PASI*************************************************************************
            'ARCV 15-02-2007
            'ALPA 20080926**************************************************************************
            'clsprevio.PrintSpool sLpt, oImpresora.gPrnTpoLetraSansSerif1PDef & oImpresora.gPrnTamLetra10CPIDef & sImpreDocs, False, gnLinPage
            '***************************************************************************************
            '*** PEAC 20080723*
            'Prev.Show oImpresora.gPrnCondensadaON & sCadImp & oImpresora.gPrnCondensadaOFF, ""
            Prev.Show oImpresora.gPrnTpoLetraSansSerif1PDef & sCadImp & oImpresora.gPrnTamLetra10CPIDef, ""
            'desproteger para compilacion de fin de mes de julio PEAC 20080730
            'Prev.Show oImpresora.gPrnTpoLetraSansSerif1PDef & oImpresora.gPrnTamLetra10CPIDef & sCadImp, ""
            
            '----
            Set Prev = Nothing
            Set oCredDoc = Nothing
            
         Case "006" 'Pagare

            '*** PEAC 20091211
            If lnEstCred < 2001 Then
                MsgBox "El crédito debe estar por lo menos en estado Sugerido para mostrar el reporte.", vbOKOnly, "Mensaje"
                Exit Sub
            End If
            '***FIN PEAC

            Set oCredDoc = New COMNCredito.NCOMCredDoc
            Set Prev = New clsprevio
            
            If vbYes = MsgBox("Formato Nuevo [SI]  / Formato Antiguo [NO]", vbInformation + vbYesNo) Then
                nTipoFormato = 0
            Else
                nTipoFormato = 1
            End If
            
            'Call ImprimePagareCred(ActXCta.NroCuenta, nTipoFormato)'WIOR 20130423 Comentado
            
            '***********************************************************
            'Modificado por PASI TI-ERS136-2013
            'Call ImprimePagareCredPDF(ActxCta.NroCuenta, nTipoFormato) 'WIOR 20130423
            bValor = VerificarExisteDesembolsoBcoNac(ActxCta.NroCuenta, sFecDes, 2)
            If bValor = True Then
                Call ImprimePagareCredPDF(ActxCta.NroCuenta, nTipoFormato, sFecDes)
            Else
                Call ImprimePagareCredPDF(ActxCta.NroCuenta, nTipoFormato)
            End If
            'END PASI
            
                                                
            'Prev.Show sCadImp, "", False
            Set Prev = Nothing
            Set oCredDoc = Nothing
            
          Case "007" 'Hoja de Reusmen
            'Set Prev = New clsPrevio
            'Prev.Show ImprimeCartillaCred(ActxCta.NroCuenta), "", False
            Call ImprimeCartillaCred(ActxCta.NroCuenta)
            'Set Prev = Nothing
            'Set oCredDoc = Nothing
            
        '*** PEAC 20080412
          Case "008" ''Aprobacion de creditos
                If Trim(lblcodigo(4).Caption) = "SOLICITADO" Then
                    MsgBox "El crédito debe estar por lo menos Sugerido para mostrar el reporte.", vbOKOnly, "Mensaje"
                    Exit Sub
                End If
                'Call ImprimeAprobacionCreditos(ActxCta.NroCuenta, lcdni, lcruc)
                Call ImprimeAprobacionCreditos(ActxCta.NroCuenta, lcDNI, lcRUC)

          Case "009" 'Informe comercial (mes o comercial)
                If Trim(lblcodigo(4).Caption) = "SOLICITADO" Then
                    MsgBox "El crédito debe estar por lo menos Sugerido para mostrar el reporte.", vbOKOnly, "Mensaje"
                    Exit Sub
                End If
                
                '*** FRHU 20160823 / LUCV20160824, Según ERS004-2016
                If Not CreditoTieneFormatoEvaluacion(ActxCta.NroCuenta) Then
                    MsgBox "El crédito no tiene Formato de Evaluación", vbInformation, "Aviso"
                    Exit Sub
                End If
                '*** FIN FRHU 20160823 / Fin LUCV20160824
                
                Set oCred = New COMNCredito.NCOMCredDoc
                    Call oCred.RecuperaDatosInformeComercial01(ActxCta.NroCuenta, R)
                Set oCred = Nothing
            
                If R.EOF And R.BOF Then
                    MsgBox "No existen Datos para el Reporte...", vbInformation, "Aviso"
                    Exit Sub
                End If
            
                'If Len(Trim(R!cNumeroFuente)) > 0 Then
                    Set oCredB = New COMNCredito.NCOMCredDoc
                    Call oCredB.RecuperaDatosBalance(ActxCta.NroCuenta, RB)
                    Set oCredB = Nothing
                'End If
                
                '**********CTI3 11092018 ---Accionistas
                'Dim oDCredtpP As COMDCredito.DCOMCredito
                Dim tpoPersoneria As Integer
                Set oDCredtpP = New COMDCredito.DCOMCredito
                Set tpoPersJur = oDCredtpP.tpoPersonaNatJur(ActxCta.NroCuenta)
                Set oDCredtpP = Nothing
                
                tpoPersoneria = tpoPersJur!nPersPersoneria
                
                    Set oDCredAdcA = New COMDCredito.DCOMCredito
                        Set RDatAdcA = oDCredAdcA.RecuperaDatosAdicionales(ActxCta.NroCuenta, 1)
                    Set oDCredAdcA = Nothing
                    
                        If RDatAdcA.RecordCount <= 0 And tpoPersoneria > 1 Then
                            MsgBox "Se necesita actualizar los Datos Accionariales, ingrese a la Ruta Personas > Mantenimiento y proceda a Actualizar.", vbInformation, "AVISO"
                            Screen.MousePointer = 0
                           Exit Sub
                        Else
                            If tpoPersoneria = 1 Then
                                MsgBox "Recuerde: Si el Cliente cuenta con Datos Accionariales, ingrese a la ruta: Personas > Mantenimiento y proceda a Actualizar.", vbInformation, "AVISO"
                                'Screen.MousePointer = 0
                            End If
                        End If
                    'directorio
                    Set oDCredAdcD = New COMDCredito.DCOMCredito
                        Set RDatAdcD = oDCredAdcD.RecuperaDatosAdicionales(ActxCta.NroCuenta, 2)
                    Set oDCredAdcD = Nothing
                    'gerencias
                    Set oDCredAdcG = New COMDCredito.DCOMCredito
                        Set RDatAdcG = oDCredAdcG.RecuperaDatosAdicionales(ActxCta.NroCuenta, 3)
                    Set oDCredAdcG = Nothing
                    'patrimonio
                    Set oDCredAdcP = New COMDCredito.DCOMCredito
                        Set RDatAdcP = oDCredAdcP.RecuperaDatosAdicionales(ActxCta.NroCuenta, 4)
                    Set oDCredAdcP = Nothing
                    'cargos
                    Set oDCredAdcCa = New COMDCredito.DCOMCredito
                        Set RDatAdcCa = oDCredAdcCa.RecuperaDatosAdicionales(ActxCta.NroCuenta, 6)
                    Set oDCredAdcCa = Nothing
                    '*************************
                        
                Set oDCred = New COMDCredito.DCOMCredito
                    Set RDatFin = oDCred.RecuperaDatosFinan(ActxCta.NroCuenta)
                Set oDCred = Nothing
                    
                Call ImprimeInformeComercial01(ActxCta.NroCuenta, gsNomAge, gsCodUser, R, RB, RDatFin, RDatAdcA, RDatAdcD, RDatAdcG, RDatAdcP, RDatAdcCa) '****** CTI3 12092018

          Case "010" 'Informe Comercial (consumo)
                If Trim(lblcodigo(4).Caption) = "SOLICITADO" Then
                    MsgBox "El crédito debe estar por lo menos Sugerido para mostrar el reporte.", vbOKOnly, "Mensaje"
                    Exit Sub
                End If
                
                '*** FRHU 20160823 / LUCV20160824, Según ERS004-2016
                If Not CreditoTieneFormatoEvaluacion(ActxCta.NroCuenta) Then
                    MsgBox "El crédito no tiene Formato de Evaluación", vbInformation, "Aviso"
                    Exit Sub
                End If
                '*** FIN FRHU 20160823 / Fin LUCV20160824
                
                Set oCred = New COMNCredito.NCOMCredDoc
                Call oCred.RecuperaDatosInformeComercial(ActxCta.NroCuenta, R)
                Set oCred = Nothing
            
                If R.EOF And R.BOF Then
                    MsgBox "No existen Datos para el Reporte...", vbInformation, "Aviso"
                    Exit Sub
                End If
            
                'If Len(Trim(R!cNumeroFuente)) > 0 Then
                    Set oCredB = New COMNCredito.NCOMCredDoc
                    Call oCredB.RecuperaDatosBalance(ActxCta.NroCuenta, RB)
                    Set oCredB = Nothing
                'End If

                Set oDCredC = New COMDCredito.DCOMCredito
                    Set RDatFin = oDCredC.RecuperaDatosFinan(ActxCta.NroCuenta)
                Set oDCredC = Nothing
                
               '**********CTI3 08112018
                MsgBox "Recuerde: Si el Cliente cuenta con Datos de Partipación Patrimonial en Empresas, ingrese a la ruta: Personas > Mantenimiento y proceda a Actualizar.", vbInformation, "AVISO"
                '**********CTI3 11092018
                Set oDCredAdc = New COMDCredito.DCOMCredito
                    Set RDatAdc = oDCredAdc.RecuperaDatosAdicionales(ActxCta.NroCuenta, 5)
                Set oDCredAdc = Nothing
                
                
                Set oDCredito = New COMDCredito.DCOMCredDoc
                'Set rsDatosDDJJ = oDCredito.ObtenerDatosDDJJPatri(ActxCta.NroCuenta)
                Set rsGarantTitular = oDCredito.ObtenerGarantiasTitular(ActxCta.NroCuenta)
                'Set rsGarantAvales = oDCredito.ObtenerGarantiasAvales(ActxCta.NroCuenta)
                Set oDCredito = Nothing
                                
                '*************************
                'Call ImprimeInformeComercial02(ActxCta.NroCuenta, gsNomAge, gsCodUser, R, RB, RDatFin)
                Call ImprimeInformeComercial02(ActxCta.NroCuenta, gsNomAge, gsCodUser, R, RB, RDatFin, RDatAdc, rsGarantTitular) '****CTI3    11092018

          Case "011" 'Informe de visita al cliente
               If Trim(lblcodigo(4).Caption) = "SOLICITADO" Then
                    MsgBox "El crédito debe estar por lo menos Sugerido para mostrar el reporte.", vbOKOnly, "Mensaje"
                    Exit Sub
                End If
                '*** FRHU 20160823 / LUCV20160824, Según ERS004-2016
                If Not CreditoTieneFormatoEvaluacion(ActxCta.NroCuenta) Then
                    MsgBox "El crédito no tiene Formato de Evaluación", vbInformation, "Aviso"
                    Exit Sub
                End If
                '*** FIN FRHU 20160823 / Fin LUCV20160824
                
                Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
                Set rsInfVisita = New ADODB.Recordset
                Set rsInfVisita = oDCOMFormatosEval.RecuperarDatosInformeVisitaFormato1_6(ActxCta.NroCuenta)
                           
                If (rsInfVisita.EOF And rsInfVisita.BOF) Then
                    Set oDCOMFormatosEval = Nothing
                    MsgBox "No existen datos para este reporte.", vbOKOnly, "Atención"
                    Exit Sub
                End If
                Call CargaInformeVisitaPDF(rsInfVisita)
                
               '->***** Comentado por LUCV20160812, Según ERS004-2016
                'Set oCred = New COMNCredito.NCOMCredDoc
                    'Call oCred.RecuperaDatosInformeComercial(ActxCta.NroCuenta, R)
                'Set oCred = Nothing
                
                'If R.EOF And R.BOF Then
                    'MsgBox "No existen Datos para el Reporte...", vbInformation, "Aviso"
                    'Exit Sub
                'End If
                'Call ImprimeInformeVisitaCliente(ActxCta.NroCuenta, gsNomAge, gsCodUser, R)
                'Call CargaInformeVisitaPDF
                '<-***** Fin Comentario LUCV
          Case "012" 'Criterios de aceptacion de riesgo
                 If Trim(lblcodigo(4).Caption) = "SOLICITADO" Then
                    MsgBox "El crédito debe estar por lo menos Sugerido para mostrar el reporte.", vbOKOnly, "Mensaje"
                    Exit Sub
                End If
                '*** FRHU 20160823 / LUCV20160824, Según ERS004-2016
                If Not CreditoTieneFormatoEvaluacion(ActxCta.NroCuenta) Then
                    MsgBox "El crédito no tiene Formato de Evaluación", vbInformation, "Aviso"
                    Exit Sub
                End If
                '*** FIN FRHU 20160823 / Fin LUCV20160824
                
                '->***** LUCV20160812, Comentó Según ERS004-2016
                'Set oCred = New COMNCredito.NCOMCredDoc
                'Call oCred.RecuperaDatosInformeComercial(ActxCta.NroCuenta, R)
                'Set oCred = Nothing
                'If R.EOF And R.BOF Then
                '    MsgBox "No existen Datos para el Reporte...", vbInformation, "Aviso"
                '    Exit Sub
                'End If
                'Set oDCredSbs = New COMDCredito.DCOMCredito
                    'Set RSbs = oDCredSbs.RecuperaCaliSbs(lcDNI, lcRUC) x madm 20110107
                '   Set RSbs = oDCredSbs.RecuperaCaliSbs(lcDNI, lcRUC, lsPersTDoc)
                '    Set RDatFin1 = oDCredSbs.RecuperaDatosFinan(ActxCta.NroCuenta)
                '    Set RCap = oDCredSbs.RecuperaCapacidadPago(ActxCta.NroCuenta)
                'Set oDCredSbs = Nothing
                'Call ImprimeInformeCriteriosAceptacionRiesgo(ActxCta.NroCuenta, gsNomAge, gsCodUser, R, RSbs, RDatFin1, RCap)
                '<-***** LUCV20160812 - Fin
                
                '->***** LUCV20160812, Según: ERS004-2016
                Set oNCOMFormatosEval = New COMNCredito.NCOMFormatosEval
                Call oNCOMFormatosEval.RecuperaDatosInformeComercial(ActxCta.NroCuenta, R)
                Set oNCOMFormatosEval = Nothing
                
                If R.EOF And R.BOF Then
                    MsgBox "No existen Datos para el Reporte...", vbInformation, "Aviso"
                    Exit Sub
                End If
                
                lcDNI = Trim(R!dni_deudor)
                lcRUC = Trim(R!ruc_deudor)
                
                Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
                Set RSbs = oDCOMFormatosEval.RecuperaCaliSbs(lcDNI, lcRUC)
                Set RDatFin1 = oDCOMFormatosEval.RecuperaDatosFinan(ActxCta.NroCuenta)
                Set oDCOMFormatosEval = Nothing
                'INICIO EAAS201181120 SEGUN ERS-072-2018
                Set oScore = New COMDCredito.DCOMCredito
                If Not rsRelFinales.BOF And Not rsRelFinales.EOF Then
                    For iCR = 0 To rsRelFinales.RecordCount - 1
                        cParamExtraJSON = ""
                        cParamExtraXML = ""
                        bPasarParamExperian = False
                        If Trim(rsRelFinales!cParamExtraJSON) <> "" Then
                            cParamExtraJSON = Trim(rsRelFinales!cParamExtraJSON)
                        End If

                        If Trim(rsRelFinales!cParamExtraXML) <> "" Then
                            cParamExtraXML = Trim(rsRelFinales!cParamExtraXML)
                        End If

                        If cParamExtraXML <> "" And cParamExtraJSON <> "" Then
                            bPasarParamExperian = True
                        End If

                            Set rsInforme = oScore.obtenerInforme(CStr(rsRelFinales!cPersIDTpo), rsRelFinales!cPersIDnro, rsRelFinales!ApellidoPaterno, rsRelFinales!cperscod, bPasarParamExperian, cParamExtraXML)
                        rsRelFinales.MoveNext
                    Next iCR
                End If
                nProtestosSinAclarar = 1
                Set RProtestosSinAclarar = oScore.GetProtestosSinAclarar(ActxCta.NroCuenta)
                If Not RProtestosSinAclarar.BOF And Not RProtestosSinAclarar.EOF Then
                    If (RProtestosSinAclarar!cCodigosTitular <> "") Then
                    nProtestosSinAclarar = 0
                    End If
                    Set RProtestosSinAclarar = Nothing
                End If
                cCodigosTitular = ActxCta.NroCuenta & CStr(rsInforme!nIdInforme)
                nCobranzaCoativaSunat = 1
                nObligacionesCerradas = 1
                Set RCobranzaCoativaSunat = oScore.GetCobranzaCoactivaSunat(cCodigosTitular)
                If Not RCobranzaCoativaSunat.BOF And Not RCobranzaCoativaSunat.EOF Then
                    If (RCobranzaCoativaSunat!cCodigosTitular <> "") Then
                    nCobranzaCoativaSunat = 0
                    End If
                    If (RCobranzaCoativaSunat!cCodigosTitularTDC <> "") Then
                    nObligacionesCerradas = 0
                    End If
                    Set RCobranzaCoativaSunat = Nothing
                End If
                Call rsInforme.Close
                Set rsInforme = Nothing
                Set oScore = Nothing
                'FIN EAAS20181120 SEGUN ERS-072-2018
                Call ImprimeInformeCriteriosAceptacionRiesgoFormatoEval(ActxCta.NroCuenta, gsNomAge, gsCodUser, R, RSbs, RDatFin1, nProtestosSinAclarar, nCobranzaCoativaSunat, nObligacionesCerradas) 'EAAS20181120 SEGUN ERS-072-2018
                '<-***** Fin LUCV
          Case "013" 'Clientes relacionados
            If Trim(lblcodigo(4).Caption) = "SOLICITADO" Then
                MsgBox "El crédito debe estar por lo menos Sugerido para mostrar el reporte.", vbOKOnly, "Mensaje"
                Exit Sub
            End If

            Set oCredDoc = New COMNCredito.NCOMCredDoc
            Set Prev = New clsprevio
            Prev.Show oCredDoc.ImprimeClientesRelacionados(ActxCta.NroCuenta, gsNomAge, Format(gdFecSis, "dd/mm/yyyy"), gsCodUser, gsNomCmac, lcDNI, lcRUC), "", False
            'Prev.Show oCredDoc.ImprimeClientesRelacionadosResumen(ActxCta.NroCuenta, gsNomAge, Format(gdFecSis, "dd/mm/yyyy"), gsCodUser, gsNomCmac), "", False
            Set Prev = Nothing
            Set oCredDoc = Nothing

          Case "014" 'Declaracion Jurada Patrimonial
            If Trim(lblcodigo(4).Caption) = "SOLICITADO" Then
                MsgBox "El crédito debe estar por lo menos Sugerido para mostrar el reporte.", vbOKOnly, "Mensaje"
                Exit Sub
            End If

            Set oDAgencia = New COMDConstantes.DCOMAgencias
                Set rsAgencia = oDAgencia.RecuperaAgencias(gsCodAge)
            Set oDAgencia = Nothing
        
            Set oDCredito = New COMDCredito.DCOMCredDoc
                Set rsDatosDDJJ = oDCredito.ObtenerDatosDDJJPatri(ActxCta.NroCuenta)
                Set rsGarantTitular = oDCredito.ObtenerGarantiasTitular(ActxCta.NroCuenta)
                Set rsGarantAvales = oDCredito.ObtenerGarantiasAvales(ActxCta.NroCuenta)
            Set oDCredito = Nothing
                       
            Call ImprimeDDJJPatrimonial(rsAgencia, rsDatosDDJJ, rsGarantTitular, rsGarantAvales, gdFecSis)
            
          Case "015" 'Documentos REACTIVA PERÚ   ANGC20200518
            Call ImprimeActivatePeru(ActxCta.NroCuenta) '' impresion campaña reactiva
    End Select
    'RECO20150331************************************
    Screen.MousePointer = 0
    'RECO FIN****************************************
    Exit Sub

'ErrorCmdImprimir_Click:
        'MsgBox Err.Description, vbCritical, "Aviso"
        
End Sub

'*** PEAC 20080412
Public Sub ImprimeAprobacionCreditos(ByVal psCtaCod As String, ByVal psCodNat As String, ByVal psCodJur As String)

    Dim fs As Scripting.FileSystemObject
    Dim pRs As ADODB.Recordset
    Dim xlsAplicacion As Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim lbExisteHoja As Boolean
    Dim lsArchivo1 As String
    Dim lsArchivo2 As String
    Dim lbLibroOpen As Boolean
    Dim lsNomHoja  As String
    Dim lsMes As String

    Dim oDCred As COMDCredito.DCOMCredito
    Dim oDBCred As COMDCredito.DCOMCredDoc
    Dim oCredB As COMNCredito.NCOMCredDoc
    Dim oDGarantia As COMDCredito.DCOMGarantia
    Dim oDPersGeneral As COMDPersona.DCOMPersGeneral 'MIOL 20120815, SEGUN RQ12134
    Dim oNCred As COMNCredito.NCOMCredito 'JUEZ 20120920
    Dim oDLeasing As COMDCredito.DCOMleasing 'ALPA 20121228
    Dim oDCredExoAut As COMDCredito.DCOMNivelAprobacion 'RECO20140307 ERS174-2013
    Dim oPersVinc As COMDPersona.UCOMPersona 'APRI20170719 TI-ERS025-2017
    Dim rsValCodSobPlanM As ADODB.Recordset 'JOEP 20170919 -TIC1709190004
    
    Dim R As ADODB.Recordset
    Dim RB As ADODB.Recordset
    Dim RCaliSbs As ADODB.Recordset
    Dim RRelaCred As ADODB.Recordset, RGarantCred As ADODB.Recordset
    Dim rBancos As ADODB.Recordset, RDatFin As ADODB.Recordset, rResGarTitAva As ADODB.Recordset
    Dim REstCivConvenio As ADODB.Recordset 'MIOL 20120815, SEGUN RQ12134
    Dim RIngresoNeto As ADODB.Recordset 'MIOL 20120815, SEGUN RQ12134
    Dim RCredEval As ADODB.Recordset 'JUEZ 20120920
    Dim RLeasing As ADODB.Recordset 'ALPA 20121228
    Dim RCredAmp As ADODB.Recordset 'RECO20140307 ERS174-2013
    Dim RCalfSbsRel As ADODB.Recordset 'RECO20140307 ERS174-2013
    Dim RExoAutCred As ADODB.Recordset 'RECO20140307 ERS174-2013
    Dim RRiesgoUnico As ADODB.Recordset 'RECO20140307 ERS174-2013
    Dim RCredGarant As ADODB.Recordset 'RECO20140307 ERS174-2013
    Dim RCredResulNivApr As ADODB.Recordset 'RECO20140307 ERS174-2013
    Dim RDatosConv As ADODB.Recordset 'RECO20140307 ERS174-2013
    Dim RNivApr As ADODB.Recordset 'RECO20140307 ERS174-2013
    Dim ROpRiesgo As ADODB.Recordset 'RECO20140307 ERS174-2013
    Dim sRiesgo As String 'RECO20140307 ERS174-2013
    Dim RComentAnalis As ADODB.Recordset 'RECO20140307 ERS174-2013
    Dim sCadImp As String, i As Integer, j As Integer
    Dim lsNombreArchivo As String
    Dim lMatCabecera(1, 1) As Variant, vmone As String, vage As String
    Dim rsSobreEnd As ADODB.Recordset 'WIOR 20160621
    Dim rsSobreEndCodigos As ADODB.Recordset 'WIOR 20160621
    Dim rsAutorizaciones As ADODB.Recordset 'FRHU 20160811 Anexo-002 ERS002-2016
    Dim rsVinculadosRiesgoUnico As ADODB.Recordset 'APRI 20170719 TI-ERS025 2017
    '*** FRHU 20160823
    If Not CreditoTieneFormatoEvaluacion(psCtaCod) Then
        MsgBox "El crédito no tiene Formato de Evaluación", vbInformation, "Aviso"
        Exit Sub
    End If
    '*** FIN FRHU 20160823
    'RECO20160730 ******************************************************
    Dim oDCOMFormatosEval As New COMDCredito.DCOMFormatosEval
    Dim prsDatRatios As New ADODB.Recordset
    
    Set prsDatRatios = oDCOMFormatosEval.RecuperaDatosRatios(psCtaCod)
    'RECO FIN **********************************************************
    'Screen.MousePointer = 11
    
    'JOEP 20170919 - TIC1709190004
    Set oDCred = New COMDCredito.DCOMCredito
     
    Set rsValCodSobPlanM = oDCred.SobreEndDatosCreditoADesbloqCodigos(psCtaCod)
    If Not (rsValCodSobPlanM.EOF And rsValCodSobPlanM.BOF) Then
        If rsValCodSobPlanM!cPlanmitigacion = "" Then
            MsgBox "Este crédito presenta Códigos de Sobreendeudado. " & Chr(13) & "Cargo a desbloquear: " & IIf(Left(rsValCodSobPlanM!CodFinal, 1) = "P", "Jefe de Agencia", "Jefe de Negocios Territoriales"), vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    'JOEP 20170919 - TIC1709190004
    
    Set oDCred = New COMDCredito.DCOMCredito
        Set R = oDCred.RecuperaDatosAprobacionCreditos(psCtaCod)
        Set RCaliSbs = oDCred.RecuperaCaliSbs(psCodNat, psCodJur, lsPersTDoc)
        Set RRelaCred = oDCred.RecuperaRelacPers(psCtaCod)
        Set RDatFin = oDCred.RecuperaDatosFinan(psCtaCod)
        Set RCredEval = oDCred.RecuperaColocacCredEvalAprobacion(psCtaCod)
        
    'RECO20140307 ERS174-2013******************************
    Set oDCredExoAut = New COMDCredito.DCOMNivelAprobacion
   
    Set ROpRiesgo = oDCred.RecuperaOpinionRiesgo(psCtaCod)
    Set RCredAmp = oDCred.VerificarAmpliados(psCtaCod)
    Set RCalfSbsRel = oDCred.ObtieneCalifSBSRelacionCred(psCtaCod)
    Set RExoAutCred = oDCredExoAut.ObtieneExoneraAutoriCred(psCtaCod)
    Set RRiesgoUnico = oDCredExoAut.ObtineRiesgoUnicoCred(psCtaCod)
    Set RCredGarant = oDCred.ObtieneDatasGarantiaCred(psCtaCod)
    Set RCredResulNivApr = oDCredExoAut.ObtieneNivelAprResultado(psCtaCod)
    
    'Set RNivApr = oDCredExoAut.RecuperaHistorialCredAprobados(psCtaCod)
        
    Set RComentAnalis = oDCred.RecuperaComentarioAnalistaSugerencia(psCtaCod)
    'RECO FIN *********************************************
    Set rsSobreEnd = oDCred.SobreEndObtenerCodigosRegXCta(psCtaCod) 'WIOR 20160621
    Set rsSobreEndCodigos = oDCred.SobreEndDatosCreditoADesbloqCodigos(psCtaCod) 'WIOR 20160621
    Set rsAutorizaciones = oDCred.MostrarAutorizacionesHojaAprobacion(psCtaCod) 'FRHU 20160811 Anexo-002 ERS002-2016
    Set oDCred = Nothing
    'ALPA 20121228****************Ç
    Set oDLeasing = New COMDCredito.DCOMleasing
    Set RLeasing = oDLeasing.Obtener_MontoFinanciarLeasing(psCtaCod)
    Set oDLeasing = Nothing
    '*****************************
    'APRI 20170719 TI-ERS025 2017
     Set oPersVinc = New COMDPersona.UCOMPersona
    Set rsVinculadosRiesgoUnico = oPersVinc.ObtenerVinculadoRiesgoUnico("", psCtaCod, 0)
    Set oPersVinc = Nothing
    'END APRI
    If R.RecordCount = 0 Then
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        'Screen.MousePointer = 0
        Exit Sub
    End If

    Set oDGarantia = New COMDCredito.DCOMGarantia
        Set RGarantCred = oDGarantia.RecuperaGarantiaCredito(psCtaCod)
    Set oDGarantia = Nothing


    Set oDBCred = New COMDCredito.DCOMCredDoc
        'Set rBancos = oDBCred.RecuperaRelaBancosPersona(psCodNat, psCodJur, psCtaCod)
        Set rBancos = oDBCred.RecuperaRelaBancosPersonaHojaApr(psCodNat, psCodJur, psCtaCod)
        Set rResGarTitAva = oDBCred.RecuperaResumenGarTitularAval(psCtaCod)
    Set oDBCred = Nothing

  'Determinando que Archivo y hoja Excel se debe abrir de acuerdo a eleccion del usuario
    'MIOL 20120813, SEGUN RQ12134 ****************************************************
    'If Not (Mid(psCtaCod, 6, 3) = "515" Or Mid(psCtaCod, 6, 3) = "516" Or Mid(psCtaCod, 6, 3) = "704") Then
    '    lsArchivo1 = "Aprobacion_de_Creditos"
    'ElseIf Mid(psCtaCod, 6, 3) = "704" Then
    '    lsArchivo1 = "AprobacionxConvenio"
    'Else
    '    lsArchivo1 = "Aprobacion_AF"
    'End If
    'END MIOL ************************************************************************
    'lsNomHoja = "AproCred"
    
    'Set fs = New Scripting.FileSystemObject
    'Set xlsAplicacion = New Excel.Application

    'If fs.FileExists(App.path & "\FormatoCarta\" & lsArchivo1 & ".xls") Then
    '    Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo1 & ".xls")
    'Else
    '    MsgBox "No Existe Plantilla en Carpeta FormatoCarta, Consulte con el Area de  TI", vbInformation, "Advertencia"
    '    Exit Sub
    'End If

    'lsArchivo2 = lsArchivo1 & "_" & gsCodUser & "_" & Format$(gdFecSis, "yyyymmdd") & "_" & Format$(Time(), "HHMMSS")

    'For Each xlHoja1 In xlsLibro.Worksheets
    '   If xlHoja1.Name = lsNomHoja Then
    '        xlHoja1.Activate
    '     lbExisteHoja = True
    '    Exit For
    '   End If
    'Next
    'If lbExisteHoja = False Then
    '    Set xlHoja1 = xlsLibro.Worksheets
    '    xlHoja1.Name = lsNomHoja
    'End If
   'MIOL 20120814, SEGUN RQ12134 *******************************************************************
    
'     If Mid(psCtaCod, 6, 3) = "704" Then
'        'Genera Estado Civil y Institucion de convenio
'        Set oDPersGeneral = New COMDPersona.DCOMPersGeneral
'                    Set REstCivConvenio = oDPersGeneral.RecuperaEstCivConvenio(psCtaCod, Format$(gdFecSis, "yyyymmdd"))
'        Set oDPersGeneral = Nothing
'        'Genera el Ingreso Neto
'        Set oDPersGeneral = New COMDPersona.DCOMPersGeneral
'            Set RIngresoNeto = oDPersGeneral.RecuperaIngresoNeto(psCtaCod)
'        Set oDPersGeneral = Nothing
'        'Call InsertaDatosAprobCredxConvenio(R, RCaliSbs, RRelaCred, rBancos, RDatFin, xlHoja1, rResGarTitAva, RGarantCred, REstCivConvenio, RIngresoNeto)
'        Call ImprimeHojaAprovacionCredConsumo(R, RCaliSbs, RRelaCred, rBancos, RDatFin, rResGarTitAva, RGarantCred, psCtaCod, RCredEval, RCredAmp, RCalfSbsRel, RExoAutCred, RRiesgoUnico, RCredGarant, RCredResulNivApr, REstCivConvenio)
'    Else
'        'Call InsertaDatosAprobCred(R, RCaliSbs, RRelaCred, rBancos, RDatFin, xlHoja1, rResGarTitAva, RGarantCred, RCredEval, RLeasing)
'        Call ImprimeHojaAprovacionCred(R, RCaliSbs, RRelaCred, rBancos, RDatFin, rResGarTitAva, RGarantCred, psCtaCod, RCredEval, RCredAmp, RCalfSbsRel, RExoAutCred, RRiesgoUnico, RCredGarant, RCredResulNivApr)
'    End If
    'END MIOL **************************************************************************************
    'RECO20143012 ERS174-2013***********************************************************************
    Set oDPersGeneral = New COMDPersona.DCOMPersGeneral
        Set REstCivConvenio = oDPersGeneral.RecuperaEstCivConvenio(psCtaCod, Format$(gdFecSis, "yyyymmdd"))
    Set oDPersGeneral = Nothing
    If val(Mid(psCtaCod, 6, 3)) >= 800 And val(Mid(psCtaCod, 6, 3)) < 900 Then
        sRiesgo = "RIESGO 2"
    Else
        sRiesgo = "RIESGO 1"
    End If
    '
    'RECO20140307 ERS174-2013******************************
    Set oDCredExoAut = New COMDCredito.DCOMNivelAprobacion   '***CTI3 (ferimoro) ERS062-2018
    
    
    Dim oConst As COMDConstSistema.DCOMGeneral
    Dim sConstante As String
    Set oConst = New COMDConstSistema.DCOMGeneral
    sConstante = oConst.LeeConstSistema(473)
    Set oConst = Nothing
    'RECO20140804*************************************
    Dim oGeneral As COMDConstSistema.DCOMGeneral 'DGeneral
    Dim nTipoCambioFijo
    Set oGeneral = New COMDConstSistema.DCOMGeneral
    nTipoCambioFijo = oGeneral.EmiteTipoCambio(gdFecSis, TCFijoMes)
    Set oGeneral = Nothing
    'RECO FIN*****************************************
    
    '**ARLO20180712 ERS042 - 2018
    Set objProducto = New COMDCredito.DCOMCredito
    If objProducto.GetResultadoCondicionCatalogo("N0000094", Mid(psCtaCod, 6, 3)) Then
    'If val(Mid(psCtaCod, 6, 3)) = 704 Then 'And val(Mid(psCtaCod, 6, 3)) < 800 Then
    '**ARLO20180712 ERS042 - 2018
        'Call ImprimeHojaAprobacionCredConsumo(R, RCaliSbs, RRelaCred, rBancos, RDatFin, rResGarTitAva, RGarantCred, psCtaCod, RCredEval, RCredAmp, RCalfSbsRel, RExoAutCred, RRiesgoUnico, RCredGarant, RCredResulNivApr, REstCivConvenio, RNivApr, sRiesgo, sConstante, ROpRiesgo, RComentAnalis, nTipoCambioFijo, rsSobreEnd, rsSobreEndCodigos, prsDatRatios)
        Call ImprimeHojaAprobacionCredConsumo(R, RCaliSbs, RRelaCred, rBancos, RDatFin, rResGarTitAva, RGarantCred, psCtaCod, RCredEval, RCredAmp, RCalfSbsRel, RExoAutCred, RRiesgoUnico, RCredGarant, RCredResulNivApr, REstCivConvenio, RNivApr, sRiesgo, sConstante, ROpRiesgo, RComentAnalis, nTipoCambioFijo, rsSobreEnd, rsSobreEndCodigos, prsDatRatios, rsAutorizaciones, rsVinculadosRiesgoUnico) 'FRHU 20160811 Anexo002 ERS002-2016 'APRI20170719 TI-ERS025-2017 add rsVinculadosRiesgoUnico
    '**ARLO20180712 ERS042 - 2018
    ElseIf objProducto.GetResultadoCondicionCatalogo("N0000095", Mid(psCtaCod, 6, 3)) Then
    'ElseIf Mid(psCtaCod, 6, 3) = "515" Or Mid(psCtaCod, 6, 3) = "516" Then
    '**ARLO20180712 ERS042 - 2018
        'Call ImprimeHojaAprobacionCredLeasing(R, RCaliSbs, RRelaCred, rBancos, RDatFin, rResGarTitAva, RGarantCred, psCtaCod, RCredEval, RCredAmp, RCalfSbsRel, RExoAutCred, RRiesgoUnico, RCredGarant, RCredResulNivApr, RNivApr, sRiesgo, sConstante, ROpRiesgo, RComentAnalis, nTipoCambioFijo, rsSobreEnd, rsSobreEndCodigos, prsDatRatios)
        Call ImprimeHojaAprobacionCredLeasing(R, RCaliSbs, RRelaCred, rBancos, RDatFin, rResGarTitAva, RGarantCred, psCtaCod, RCredEval, RCredAmp, RCalfSbsRel, RExoAutCred, RRiesgoUnico, RCredGarant, RCredResulNivApr, RNivApr, sRiesgo, sConstante, ROpRiesgo, RComentAnalis, nTipoCambioFijo, rsSobreEnd, rsSobreEndCodigos, prsDatRatios, rsAutorizaciones, rsVinculadosRiesgoUnico) 'FRHU 20160811 Anexo002 ERS002-2016 'APRI20170719 TI-ERS025-2017 add rsVinculadosRiesgoUnico
    Else
        'Call ImprimeHojaAprobacionCred(R, RCaliSbs, RRelaCred, rBancos, RDatFin, rResGarTitAva, RGarantCred, psCtaCod, RCredEval, RCredAmp, RCalfSbsRel, RExoAutCred, RRiesgoUnico, RCredGarant, RCredResulNivApr, RNivApr, sRiesgo, sConstante, ROpRiesgo, RComentAnalis, nTipoCambioFijo, rsSobreEnd, rsSobreEndCodigos, prsDatRatios)
        Call ImprimeHojaAprobacionCred(R, RCaliSbs, RRelaCred, rBancos, RDatFin, rResGarTitAva, RGarantCred, psCtaCod, RCredEval, RCredAmp, RCalfSbsRel, RExoAutCred, RRiesgoUnico, RCredGarant, RCredResulNivApr, RNivApr, sRiesgo, sConstante, ROpRiesgo, RComentAnalis, nTipoCambioFijo, rsSobreEnd, rsSobreEndCodigos, prsDatRatios, rsAutorizaciones, rsVinculadosRiesgoUnico) 'FRHU 20160811 Anexo002 ERS002-2016 'APRI20170719 TI-ERS025-2017 add rsVinculadosRiesgoUnico
    End If
    
    'RECO FIN***************************************************************************************
    
    'xlHoja1.SaveAs App.path & "\Spooler\" & lsArchivo2 & ".xls"
    'xlsAplicacion.Visible = True
    'xlsAplicacion.Windows(1).Visible = True
    'Set xlsAplicacion = Nothing
    'Set xlsLibro = Nothing
    'Set xlHoja1 = Nothing

End Sub

'*** PEAC 20080412

Public Sub InsertaDatosAprobCred(ByRef pR As ADODB.Recordset, ByRef pRB As ADODB.Recordset, ByRef pRCliRela As ADODB.Recordset, ByRef pRRelaBcos As ADODB.Recordset, ByRef pRDatFinan As ADODB.Recordset, ByRef xlHoja1 As Excel.Worksheet, ByRef pRResGarTitAva As ADODB.Recordset, ByRef pRGarantCred As ADODB.Recordset, _
                                ByVal prsCredEval As ADODB.Recordset, ByVal prsLeasing As ADODB.Recordset)
                                'JUEZ 20120920 Se agrego prsCredEval

Dim i As Integer
Dim lnliquidez As Double, lnCapacidadPago As Double, lnExcedente As Double
Dim lnPatriEmpre As Double, lnPatrimonio As Double, lnIngresoNeto As Double
Dim lnRentabPatrimonial As Double, lnEndeudamiento As Double
'MADM 20091119
Dim lcTpoGarantia As String
Dim lnPorGravar As Double, lnGravado As Double

Dim bTipoJustificacion As Boolean 'WIOR 20120329

 
 lcTpoGarantia = ""
 lnPorGravar = 0
 lnGravado = 0
'END MADM

Dim lnMontoFinanciarLeasing As Currency

If Not (prsLeasing.BOF Or prsLeasing.EOF) Then
    lnMontoFinanciarLeasing = IIf(IsNull(prsLeasing!nMontoFinanciar), 0, prsLeasing!nMontoFinanciar)
    xlHoja1.Cells(17, 3) = Format(lnMontoFinanciarLeasing, "#,##0.00") ' "#,##0.00"
End If

If pRDatFinan.RecordCount > 0 Then
    If pRDatFinan!PasivoCteEmp = 0 Then
        lnliquidez = 0
    Else
        lnliquidez = (pRDatFinan!activoCteEmp + pRDatFinan!InventarioEmp) / pRDatFinan!PasivoCteEmp
    End If
    
    If pRDatFinan!ActivoTotal = 0 Or pRDatFinan!PatrimonioTotal = 0 Then
        lnEndeudamiento = 0
    Else
        lnEndeudamiento = (pRDatFinan!PasivoTotal - pRDatFinan!PasCteCmacMay - pRDatFinan!PasNoCteCmacMay + pRDatFinan!ExposiCred) / pRDatFinan!PatrimonioTotal
    End If
    
    If pRDatFinan!PatrimonioEmpre = 0 Then
        lnRentabPatrimonial = 0
    Else
        lnRentabPatrimonial = pRDatFinan!IngresoNeto / pRDatFinan!PatrimonioEmpre
    End If

    If pRDatFinan!Excedente = 0 Then
        lnCapacidadPago = 0
    Else
        lnCapacidadPago = pRDatFinan!ValorCuota / pRDatFinan!Excedente
    End If
Else
    lnliquidez = 0
    lnEndeudamiento = 0
    lnRentabPatrimonial = 0
    lnCapacidadPago = 0
End If

'JUEZ 20120920 *******************************
If Not prsCredEval.EOF Then
    xlHoja1.Cells(43, 5) = IIf(prsCredEval!nPersPersoneria = 1, prsCredEval!nLiqCorrienteNat, prsCredEval!nLiqCorrienteJur)
    xlHoja1.Cells(44, 5) = CDbl(prsCredEval!nEndeudTotal) / 100
    xlHoja1.Cells(45, 5) = 0
    xlHoja1.Cells(46, 5) = (IIf(prsCredEval!nFormato = 3, prsCredEval!nCapPagoEmp, prsCredEval!nCuotaUNM)) / 100
Else
    xlHoja1.Cells(43, 5) = lnliquidez
    xlHoja1.Cells(44, 5) = lnEndeudamiento
    xlHoja1.Cells(45, 5) = lnRentabPatrimonial
    xlHoja1.Cells(46, 5) = lnCapacidadPago
End If
'END JUEZ ************************************

'MADM 20091119
If pRGarantCred.RecordCount > 0 Then
    pRGarantCred.MoveFirst
    Do While Not pRGarantCred.EOF
            lcTpoGarantia = pRGarantCred!cTpoGarantia & " / " & lcTpoGarantia
            lnPorGravar = pRGarantCred!nPorGravar + lnPorGravar
            lnGravado = pRGarantCred!nGravado + lnGravado
            pRGarantCred.MoveNext
    Loop

     If (lnPorGravar + lnGravado) > 0 Then
            xlHoja1.Cells(31, 3) = Mid(lcTpoGarantia, 1, Len(lcTpoGarantia) - 2)
            xlHoja1.Cells(31, 5) = lnPorGravar
            xlHoja1.Cells(31, 7) = lnGravado
    End If

Else
    xlHoja1.Cells(31, 3) = " "
    xlHoja1.Cells(31, 5) = 0
    xlHoja1.Cells(31, 7) = 0
End If
'END MADM

'If pRGarantCred.RecordCount > 0 Then
'    xlHoja1.Cells(31, 3) = pRGarantCred!cTpoGarantia
'    xlHoja1.Cells(31, 5) = pRGarantCred!nPorGravar
'    xlHoja1.Cells(31, 7) = pRGarantCred!nGravado
'Else
'    xlHoja1.Cells(31, 3) = " "
'    xlHoja1.Cells(31, 5) = 0
'    xlHoja1.Cells(31, 7) = 0
'End If

'JUEZ 20120920 *******************************
If Not prsCredEval.EOF Then
    xlHoja1.Cells(43, 8) = prsCredEval!nPatrimonio
    xlHoja1.Cells(44, 8) = prsCredEval!nInventario
    xlHoja1.Cells(45, 8) = IIf(prsCredEval!nFormato = 3, prsCredEval!nIngresoNeto, prsCredEval!nExcedenteFam)
    xlHoja1.Cells(46, 8) = prsCredEval!nCuotaMostrar
    If pRDatFinan.RecordCount > 0 Then
        xlHoja1.Cells(11, 4) = pRDatFinan!ExposiCred
    End If
Else
    If pRDatFinan.RecordCount > 0 Then
        xlHoja1.Cells(43, 8) = pRDatFinan!PatrimonioTotal
        xlHoja1.Cells(44, 8) = pRDatFinan!INVENTARIOtot
        xlHoja1.Cells(45, 8) = pRDatFinan!Excedente
        xlHoja1.Cells(46, 8) = pRDatFinan!ValorCuota
        
        xlHoja1.Cells(11, 4) = pRDatFinan!ExposiCred
    Else
        xlHoja1.Cells(43, 8) = 0
        xlHoja1.Cells(44, 8) = 0
        xlHoja1.Cells(45, 8) = 0
        xlHoja1.Cells(46, 8) = 0
    End If
End If
'END JUEZ ************************************

If pRResGarTitAva.RecordCount > 0 Then
    Do While Not pRResGarTitAva.EOF
        If pRResGarTitAva!TitAval = 1 Then
            xlHoja1.Cells(38, 4) = pRResGarTitAva!garpresol
            xlHoja1.Cells(39, 4) = pRResGarTitAva!garpreAsol
            xlHoja1.Cells(40, 4) = pRResGarTitAva!garNopresol
            
            xlHoja1.Cells(38, 5) = pRResGarTitAva!garpreDol
            xlHoja1.Cells(39, 5) = pRResGarTitAva!garpreADol
            xlHoja1.Cells(40, 5) = pRResGarTitAva!garNopreDol
        Else
            xlHoja1.Cells(38, 6) = pRResGarTitAva!garpresol
            xlHoja1.Cells(39, 6) = pRResGarTitAva!garpreAsol
            xlHoja1.Cells(40, 6) = pRResGarTitAva!garNopresol
            
            xlHoja1.Cells(38, 7) = pRResGarTitAva!garpreDol
            xlHoja1.Cells(39, 7) = pRResGarTitAva!garpreADol
            xlHoja1.Cells(40, 7) = pRResGarTitAva!garNopreDol
        End If
        pRResGarTitAva.MoveNext
    Loop
Else
    xlHoja1.Cells(38, 4) = 0
    xlHoja1.Cells(39, 4) = 0
    xlHoja1.Cells(40, 4) = 0
    
    xlHoja1.Cells(38, 5) = 0
    xlHoja1.Cells(39, 5) = 0
    xlHoja1.Cells(40, 5) = 0
    
    xlHoja1.Cells(38, 6) = 0
    xlHoja1.Cells(39, 6) = 0
    xlHoja1.Cells(40, 6) = 0
    
    xlHoja1.Cells(38, 7) = 0
    xlHoja1.Cells(39, 7) = 0
    xlHoja1.Cells(40, 7) = 0
End If

If pR.RecordCount > 0 Then
    xlHoja1.Cells(18, 3) = pR!PeriodoGracia
    xlHoja1.Cells(28, 3) = pR!TEMAnterior
    'JUEZ 20120920 ****************************************************************************
    If Not prsCredEval.EOF Then
        xlHoja1.Cells(48, 3) = IIf(prsCredEval!nFormato = 3, prsCredEval!nIngresoProm, prsCredEval!nVentaPromMes)
        xlHoja1.Cells(48, 5) = prsCredEval!nTotalGastoFam
        xlHoja1.Cells(48, 7) = IIf(prsCredEval!nFormato = 3, prsCredEval!nEgresoProm, prsCredEval!nCostoTotal)
        xlHoja1.Cells(48, 10) = prsCredEval!nTotalOtrosIng
    Else
        xlHoja1.Cells(48, 3) = pR!IngreVtas
        xlHoja1.Cells(48, 5) = pR!GtosFami
        xlHoja1.Cells(48, 7) = pR!CostoTot
        xlHoja1.Cells(48, 10) = pR!ConFamIng
    End If
    'END JUEZ *********************************************************************************
    xlHoja1.Cells(4, 4) = pR!Prestatario
    xlHoja1.Cells(5, 4) = "'" & IIf(Len(pR!DniRuc) = 0, "---", pR!DniRuc) & " / " & IIf(Len(pR!Ruc) = 0, "---", pR!Ruc)
    xlHoja1.Cells(6, 4) = pR!dire_domicilio
    xlHoja1.Cells(7, 4) = pR!dire_trabajo
    xlHoja1.Cells(8, 4) = pR!ExpoAntMax
    xlHoja1.Cells(13, 4) = pR!Linea_Cred
    
    'JUEZ 20120920 ****************************************************************************
    If Not prsCredEval.EOF Then
        xlHoja1.Cells(16, 6) = Trim(pR!CIIU) & " - " & Trim(pR!ActiGiro) & " - " & Trim(pR!ActiComple) & " - " & prsCredEval!cGiroNeg
    Else
        xlHoja1.Cells(16, 6) = Trim(pR!CIIU) & " - " & Trim(pR!ActiGiro) & " - " & Trim(pR!ActiComple) & " - "
    End If
    'END JUEZ *********************************************************************************
    xlHoja1.Cells(24, 6) = pR!cDestino
    xlHoja1.Cells(4, 10) = Format(pR!Fec_Soli, "mm/dd/yyyy")
    xlHoja1.Cells(5, 10) = Format(pR!fec_venc, "mm/dd/yyyy")
    xlHoja1.Cells(6, 10) = "'" & pR!Nro_Credito
    xlHoja1.Cells(7, 10) = pR!Modalidad
    xlHoja1.Cells(9, 10) = pR!Analista
    xlHoja1.Cells(10, 10) = pR!Oficina
    xlHoja1.Cells(12, 10) = pR!Tipo_Cred
    
    'MAVM 20100713 ***
    xlHoja1.Cells(13, 10) = pR!Tipo_Prod
    '***
    
    xlHoja1.Cells(16, 2) = Format(pR!Ptmo_Propto, "#,##0.00") ' "#,##0.00"
    
    xlHoja1.Cells(20, 3) = Format(pR!Plazo, "#,##0.00")
    xlHoja1.Cells(21, 3) = Format(pR!PlazoEntreDias, "#,##0.00")
    xlHoja1.Cells(22, 3) = pR!Moneda
    xlHoja1.Cells(24, 3) = Format(pR!Tasa_Interes, "#,##0.00")
    xlHoja1.Cells(34, 4) = "'" & pR!registro_inmueble
    xlHoja1.Cells(34, 6) = Trim(pR!ubicacion_garant)

Else
    xlHoja1.Cells(18, 3) = 0
    xlHoja1.Cells(28, 3) = 0
    xlHoja1.Cells(48, 3) = 0
    xlHoja1.Cells(48, 5) = 0
    xlHoja1.Cells(48, 7) = 0
    xlHoja1.Cells(48, 10) = 0
    xlHoja1.Cells(4, 4) = 0
    xlHoja1.Cells(5, 4) = " "
    xlHoja1.Cells(6, 4) = " "
    xlHoja1.Cells(7, 4) = " "
    xlHoja1.Cells(8, 4) = 0
    xlHoja1.Cells(13, 4) = " "

    xlHoja1.Cells(16, 6) = " "
    xlHoja1.Cells(24, 6) = " "
    xlHoja1.Cells(4, 10) = " "
    xlHoja1.Cells(5, 10) = " "
    xlHoja1.Cells(6, 10) = " "
    xlHoja1.Cells(7, 10) = " "
    xlHoja1.Cells(9, 10) = " "
    xlHoja1.Cells(10, 10) = " "
    xlHoja1.Cells(12, 10) = " "
    xlHoja1.Cells(16, 2) = 0
    xlHoja1.Cells(20, 3) = 0
    xlHoja1.Cells(21, 3) = 0
    xlHoja1.Cells(22, 3) = " "
    xlHoja1.Cells(24, 3) = 0
    xlHoja1.Cells(34, 4) = " "
    xlHoja1.Cells(34, 6) = " "
    xlHoja1.Cells(3, 6) = ""
End If


If pRB.RecordCount > 0 Then
    xlHoja1.Cells(16, 5) = pRB!nNormal
    xlHoja1.Cells(18, 5) = pRB!nPotencial
    xlHoja1.Cells(20, 5) = pRB!nDeficiente
    xlHoja1.Cells(22, 5) = pRB!nDudoso
    xlHoja1.Cells(24, 5) = pRB!nPerdido
Else
    xlHoja1.Cells(16, 5) = 0
    xlHoja1.Cells(18, 5) = 0
    xlHoja1.Cells(20, 5) = 0
    xlHoja1.Cells(22, 5) = 0
    xlHoja1.Cells(24, 5) = 0
End If


i = 53
'*** PEAC 20080412 - CLIENTES RELACIONADOS
Do While Not pRCliRela.EOF
        i = i + 1
        'madm 20110303
        bvalorNegativo = 0
        Set oPersona = New COMDPersona.UCOMPersona
               'If oPersona.ValidaEnListaNegativaCondicion(psCodNat, psCodJur, lnCondicion, lblcodigo(1).Caption) Then
                Dim psDNI As String
                Dim psRUC As String
                psDNI = IIf(IsNull(pRCliRela!DNI), "", pRCliRela!DNI)
                psRUC = IIf(IsNull(pRCliRela!Ruc), "", pRCliRela!Ruc)
                If oPersona.ValidaEnListaNegativaCondicion(psDNI, psRUC, lnCondicion, pRCliRela!cPersNombre) Then
                    'If lnCondicion >= 1 Then Comentado by JACA 20111114
                    If lnCondicion = 1 Then ' JACA 20111114 - Solo los Negativos
                        bTipoJustificacion = oPersona.ValidaEnListaNegativaCondicionJustificacion(psDNI, psRUC, lnCondicion, pRCliRela!cPersNombre) 'WIOR 20120329
                        bvalorNegativo = 1
                    End If
                End If
        Set oPersona = Nothing
        'end madm
        
        xlHoja1.Cells(i, 2) = pRCliRela!cPersNombre
        xlHoja1.Cells(i, 5) = pRCliRela!cConsDescripcion
        'xlHoja1.Cells(i, 7) = IIf(bvalorNegativo = 1, "CBNN", "")
         xlHoja1.Cells(i, 7) = IIf(bvalorNegativo = 1, IIf(bTipoJustificacion = True, "CBNN", ""), "") 'WIOR 20120329
         
        xlHoja1.rows(i + 1).Select
        xlHoja1.Range("B" + Trim(str(i + 1))).EntireRow.Insert
        'FRHU 20131202
        If i = 54 Then
            'Set oCliPre = New COMNCredito.NCOMCredito 'COMENTADO POR ARLO 20170722
            'bValidarCliPre = oCliPre.ValidarClientePreferencial(pRCliRela!cPersCod) 'COMENTADO POR ARLO 20170722
            bValidarCliPre = False 'ARLO 20170722
            If bValidarCliPre Then
                xlHoja1.Range("B1", "K1").MergeCells = True
                xlHoja1.Range("B1", "K1").HorizontalAlignment = xlCenter
                xlHoja1.Range("B1", "K1").Font.Bold = True
                xlHoja1.Range("B1", "K1").Font.Size = 16
                xlHoja1.Cells(1, 2) = "CLIENTE PREFERENCIAL" 'pRCliRela!cPersCod
            End If
        End If
        'FIN FRHU 20131202
    pRCliRela.MoveNext
Loop
'If pRCliRela.EOF Then
'    'FRHU 20131202
'    If ValidarClientePreferencial(pRCliRela!cPersCod) Then
'        xlHoja1.Cells(1, 4) = "CLIENTE PREFERENCIAL" 'pRCliRela!cPersCod
'    End If
'    'FIN FRHU 20131202
'End If

i = i + 3
'*** PEAC 20080412 - Bancos con los que trabaja
Do While Not pRRelaBcos.EOF
        i = i + 1
        xlHoja1.Cells(i, 2) = pRRelaBcos!Nombre
        xlHoja1.Cells(i, 5) = pRRelaBcos!Moneda
        xlHoja1.Cells(i, 6) = pRRelaBcos!Saldo
        xlHoja1.Cells(i, 7) = pRRelaBcos!Relacion
        
        xlHoja1.rows(i + 1).Select
        xlHoja1.Range("B" + Trim(str(i + 1))).EntireRow.Insert

    pRRelaBcos.MoveNext
Loop

End Sub

'MIOL 20120814, SEGUN RQ12134 ********************************************************************
Public Sub InsertaDatosAprobCredxConvenio(ByRef pR As ADODB.Recordset, ByRef pRB As ADODB.Recordset, ByRef pRCliRela As ADODB.Recordset, ByRef pRRelaBcos As ADODB.Recordset, ByRef pRDatFinan As ADODB.Recordset, ByRef xlHoja1 As Excel.Worksheet, ByRef pRResGarTitAva As ADODB.Recordset, ByRef pRGarantCred As ADODB.Recordset, ByRef pREstCivConvenio As ADODB.Recordset, ByRef pRIngresoNeto As ADODB.Recordset)
Dim i As Integer
Dim lnliquidez As Double, lnCapacidadPago As Double, lnExcedente As Double
Dim lnPatriEmpre As Double, lnPatrimonio As Double, lnIngresoNeto As Double
Dim lnRentabPatrimonial As Double, lnEndeudamiento As Double
'MADM 20091119
Dim lcTpoGarantia As String
Dim lnPorGravar As Double, lnGravado As Double

Dim bTipoJustificacion As Boolean 'WIOR 20120329

 lcTpoGarantia = ""
 lnPorGravar = 0
 lnGravado = 0
'END MADM
xlHoja1.Cells(44, 3) = pRIngresoNeto!Monto
'MADM 20091119
If pRGarantCred.RecordCount > 0 Then
    pRGarantCred.MoveFirst
    Do While Not pRGarantCred.EOF
            lcTpoGarantia = pRGarantCred!cTpoGarantia & " / " & lcTpoGarantia
            lnPorGravar = pRGarantCred!nPorGravar + lnPorGravar
            lnGravado = pRGarantCred!nGravado + lnGravado
            pRGarantCred.MoveNext
    Loop
    If (lnPorGravar + lnGravado) > 0 Then
            xlHoja1.Cells(31, 3) = Mid(lcTpoGarantia, 1, Len(lcTpoGarantia) - 2)
            xlHoja1.Cells(31, 5) = lnPorGravar
            xlHoja1.Cells(31, 7) = lnGravado
    End If
Else
    xlHoja1.Cells(31, 3) = " "
    xlHoja1.Cells(31, 5) = 0
    xlHoja1.Cells(31, 7) = 0
End If
'END MADM
If pRDatFinan.RecordCount > 0 Then
    xlHoja1.Cells(11, 4) = pRDatFinan!ExposiCred
Else
    xlHoja1.Cells(11, 4) = 0
End If

If pRResGarTitAva.RecordCount > 0 Then
    Do While Not pRResGarTitAva.EOF
        If pRResGarTitAva!TitAval = 1 Then
            xlHoja1.Cells(38, 4) = pRResGarTitAva!garpresol
            xlHoja1.Cells(39, 4) = pRResGarTitAva!garpreAsol
            xlHoja1.Cells(40, 4) = pRResGarTitAva!garNopresol

            xlHoja1.Cells(38, 5) = pRResGarTitAva!garpreDol
            xlHoja1.Cells(39, 5) = pRResGarTitAva!garpreADol
            xlHoja1.Cells(40, 5) = pRResGarTitAva!garNopreDol
        Else
            xlHoja1.Cells(38, 6) = pRResGarTitAva!garpresol
            xlHoja1.Cells(39, 6) = pRResGarTitAva!garpreAsol
            xlHoja1.Cells(40, 6) = pRResGarTitAva!garNopresol

            xlHoja1.Cells(38, 7) = pRResGarTitAva!garpreDol
            xlHoja1.Cells(39, 7) = pRResGarTitAva!garpreADol
            xlHoja1.Cells(40, 7) = pRResGarTitAva!garNopreDol
        End If
        pRResGarTitAva.MoveNext
    Loop
Else
    xlHoja1.Cells(38, 4) = 0
    xlHoja1.Cells(39, 4) = 0
    xlHoja1.Cells(40, 4) = 0

    xlHoja1.Cells(38, 5) = 0
    xlHoja1.Cells(39, 5) = 0
    xlHoja1.Cells(40, 5) = 0

    xlHoja1.Cells(38, 6) = 0
    xlHoja1.Cells(39, 6) = 0
    xlHoja1.Cells(40, 6) = 0

    xlHoja1.Cells(38, 7) = 0
    xlHoja1.Cells(39, 7) = 0
    xlHoja1.Cells(40, 7) = 0
End If

If pR.RecordCount > 0 Then
    xlHoja1.Cells(18, 3) = pR!PeriodoGracia
    xlHoja1.Cells(28, 3) = pR!TEMAnterior
    xlHoja1.Cells(48, 3) = pR!IngreVtas
    xlHoja1.Cells(48, 5) = pR!GtosFami
    xlHoja1.Cells(48, 7) = pR!CostoTot
    xlHoja1.Cells(48, 10) = pR!ConFamIng
    xlHoja1.Cells(4, 4) = pR!Prestatario
    xlHoja1.Cells(5, 4) = "'" & IIf(Len(pR!DniRuc) = 0, "---", pR!DniRuc) & " / " & IIf(Len(pR!Ruc) = 0, "---", pR!Ruc)
    xlHoja1.Cells(6, 4) = pR!dire_domicilio
    xlHoja1.Cells(7, 4) = pR!dire_trabajo
    xlHoja1.Cells(8, 4) = pR!ExpoAntMax
    xlHoja1.Cells(9, 4) = pREstCivConvenio!nEdad & " - " & pREstCivConvenio!cestciv
    xlHoja1.Cells(10, 4) = pREstCivConvenio!cInstitucion
    xlHoja1.Cells(13, 4) = pR!Linea_Cred

    xlHoja1.Cells(16, 6) = Trim(pR!CIIU) & " - " & Trim(pR!ActiGiro) & " - " & Trim(pR!ActiComple)
    xlHoja1.Cells(24, 6) = pR!cDestino
    xlHoja1.Cells(4, 10) = Format(pR!Fec_Soli, "mm/dd/yyyy")
    xlHoja1.Cells(5, 10) = Format(pR!fec_venc, "mm/dd/yyyy")
    xlHoja1.Cells(6, 10) = "'" & pR!Nro_Credito
    xlHoja1.Cells(7, 10) = pR!Modalidad
    xlHoja1.Cells(8, 10) = pREstCivConvenio!nAntiguedad & " Años "
    xlHoja1.Cells(9, 10) = pR!Analista
    xlHoja1.Cells(10, 10) = pR!Oficina
    xlHoja1.Cells(12, 10) = pR!Tipo_Cred
    'MAVM 20100713 ***
    xlHoja1.Cells(13, 10) = pR!Tipo_Prod
    '***
    xlHoja1.Cells(16, 2) = Format(pR!Ptmo_Propto, "#,##0.00")
    xlHoja1.Cells(20, 3) = Format(pR!Plazo, "#,##0.00")
    xlHoja1.Cells(21, 3) = Format(pR!PlazoEntreDias, "#,##0.00")
    xlHoja1.Cells(22, 3) = pR!Moneda
    xlHoja1.Cells(24, 3) = Format(pR!Tasa_Interes, "#,##0.00")

Else
    xlHoja1.Cells(18, 3) = 0
    xlHoja1.Cells(28, 3) = 0
    xlHoja1.Cells(48, 3) = 0
    xlHoja1.Cells(48, 5) = 0
    xlHoja1.Cells(48, 7) = 0
    xlHoja1.Cells(48, 10) = 0
    xlHoja1.Cells(4, 4) = 0
    xlHoja1.Cells(5, 4) = " "
    xlHoja1.Cells(6, 4) = " "
    xlHoja1.Cells(7, 4) = " "
    xlHoja1.Cells(8, 4) = 0
    xlHoja1.Cells(13, 4) = " "

    xlHoja1.Cells(16, 6) = " "
    xlHoja1.Cells(24, 6) = " "
    xlHoja1.Cells(4, 10) = " "
    xlHoja1.Cells(5, 10) = " "
    xlHoja1.Cells(6, 10) = " "
    xlHoja1.Cells(7, 10) = " "
    xlHoja1.Cells(8, 10) = " "
    xlHoja1.Cells(9, 10) = " "
    xlHoja1.Cells(10, 10) = " "
    xlHoja1.Cells(12, 10) = " "
    xlHoja1.Cells(16, 2) = 0
    xlHoja1.Cells(20, 3) = 0
    xlHoja1.Cells(21, 3) = 0
    xlHoja1.Cells(22, 3) = " "
    xlHoja1.Cells(24, 3) = 0
    xlHoja1.Cells(34, 4) = " "
    xlHoja1.Cells(34, 6) = " "
    xlHoja1.Cells(3, 6) = ""
End If

If pRB.RecordCount > 0 Then
    xlHoja1.Cells(16, 5) = pRB!nNormal
    xlHoja1.Cells(18, 5) = pRB!nPotencial
    xlHoja1.Cells(20, 5) = pRB!nDeficiente
    xlHoja1.Cells(22, 5) = pRB!nDudoso
    xlHoja1.Cells(24, 5) = pRB!nPerdido
Else
    xlHoja1.Cells(16, 5) = 0
    xlHoja1.Cells(18, 5) = 0
    xlHoja1.Cells(20, 5) = 0
    xlHoja1.Cells(22, 5) = 0
    xlHoja1.Cells(24, 5) = 0
End If

i = 52
'*** PEAC 20080412 - CLIENTES RELACIONADOS
Do While Not pRCliRela.EOF
        i = i + 1
        'MADM 20110303
        bvalorNegativo = 0
        Set oPersona = New COMDPersona.UCOMPersona
                Dim psDNI As String
                Dim psRUC As String
                psDNI = IIf(IsNull(pRCliRela!DNI), "", pRCliRela!DNI)
                psRUC = IIf(IsNull(pRCliRela!Ruc), "", pRCliRela!Ruc)
                If oPersona.ValidaEnListaNegativaCondicion(psDNI, psRUC, lnCondicion, pRCliRela!cPersNombre) Then
                    If lnCondicion = 1 Then ' JACA 20111114 - Solo los Negativos
                        bTipoJustificacion = oPersona.ValidaEnListaNegativaCondicionJustificacion(psDNI, psRUC, lnCondicion, pRCliRela!cPersNombre) 'WIOR 20120329
                        bvalorNegativo = 1
                    End If
                End If
        Set oPersona = Nothing
        'END MADM
        xlHoja1.Cells(i, 2) = pRCliRela!cPersNombre
        xlHoja1.Cells(i, 5) = pRCliRela!cConsDescripcion
        xlHoja1.Cells(i, 7) = IIf(bvalorNegativo = 1, IIf(bTipoJustificacion = True, "CBNN", ""), "") 'WIOR 20120329
         
        xlHoja1.rows(i + 1).Select
        xlHoja1.Range("B" + Trim(str(i + 1))).EntireRow.Insert
        'FRHU 20131202
        If i = 53 Then
            'Set oCliPre = New COMNCredito.NCOMCredito 'COMENTADO POR ARLO 20170722
            'bValidarCliPre = oCliPre.ValidarClientePreferencial(pRCliRela!cPersCod) 'COMENTADO POR ARLO 20170722
            bValidarCliPre = False 'ARLO 20170722
            If bValidarCliPre Then
                xlHoja1.Range("B1", "K1").MergeCells = True
                xlHoja1.Range("B1", "K1").HorizontalAlignment = xlCenter
                xlHoja1.Range("B1", "K1").Font.Bold = True
                xlHoja1.Range("B1", "K1").Font.Size = 16
                xlHoja1.Cells(1, 2) = "CLIENTE PREFERENCIAL" 'pRCliRela!cPersCod
            End If
        End If
        'FIN FRHU 20131202
    pRCliRela.MoveNext
Loop

i = i + 4
'*** PEAC 20080412 - Bancos con los que trabaja
Do While Not pRRelaBcos.EOF
        i = i + 1
        xlHoja1.Cells(i, 2) = pRRelaBcos!Nombre
        xlHoja1.Cells(i, 5) = pRRelaBcos!Moneda
        xlHoja1.Cells(i, 6) = pRRelaBcos!Saldo
        xlHoja1.Cells(i, 7) = pRRelaBcos!Relacion
        
        xlHoja1.rows(i + 1).Select
        xlHoja1.Range("B" + Trim(str(i + 1))).EntireRow.Insert

    pRRelaBcos.MoveNext
Loop
End Sub
'END MIOL ****************************************************************************************

Private Sub CmdNewBusq_Click()
    
    lblcodigo(0).Caption = ""
    lblcodigo(4).Caption = ""
    lblcodigo(1).Caption = ""
    lblcodigo(2).Caption = ""
    lblcodigo(3).Caption = ""
    LstCred.Clear
    ActxCta.NroCuenta = ""
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 And ActxCta.Enabled = True Then 'F12
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.inicia(gColPYMEEmp, False)
        If sCuenta <> "" Then
            ActxCta.NroCuenta = sCuenta
            ActxCta.SetFocusCuenta
        End If
    End If
End Sub

Private Sub Form_Load()
Dim L As MSComctlLib.ListItem  'ListItem
    CentraForm Me
    'FRHU 20140722 ERS105-2014
    Dim oGen As COMDConstSistema.DCOMGeneral
    Set oGen = New COMDConstSistema.DCOMGeneral
    bPermisoCargo = oGen.VerificaExistePermisoCargo(gsCodCargo, PermisoCargos.gPlanDePagos)
    
    'FIN FRHU 20140722
    Set L = LstReportes.ListItems.Add(, , "Registro de Solicitud de Credito")
    L.SubItems(1) = "001"
    Set L = LstReportes.ListItems.Add(, , "Solicitud de Credito")
    L.SubItems(1) = "002"
    Set L = LstReportes.ListItems.Add(, , "Resumen de Comite")
    L.SubItems(1) = "003"
    Set L = LstReportes.ListItems.Add(, , "Comprobante de Desembolso")
    L.SubItems(1) = "004"
    If bPermisoCargo Then 'FRHU 20140722 ERS105-2014: Se agrego If bPermisoCargo Then
        Set L = LstReportes.ListItems.Add(, , "Plan de Pagos")
        L.SubItems(1) = "005"
    End If
    
    '*** PEAC 20170622
    If oGen.VerificaExistePermisoCargo(gsCodCargo, 11) Then
        Set L = LstReportes.ListItems.Add(, , "Pagare")
        L.SubItems(1) = "006"
    End If
    
    Set L = LstReportes.ListItems.Add(, , "Hoja de Resumen")
    L.SubItems(1) = "007"
    
'*** PEAC 20080412
    Set L = LstReportes.ListItems.Add(, , "Aprobación de créditos")
    L.SubItems(1) = "008"

    Set L = LstReportes.ListItems.Add(, , "Informe Comercial (Mes o Comercial)")
    L.SubItems(1) = "009"

    Set L = LstReportes.ListItems.Add(, , "Informe Comercial (Consumo)")
    L.SubItems(1) = "010"

    Set L = LstReportes.ListItems.Add(, , "Informe de Visita al Cliente")
    L.SubItems(1) = "011"

    Set L = LstReportes.ListItems.Add(, , "Criterios de Aceptación de Riesgo")
    L.SubItems(1) = "012"
    
    Set L = LstReportes.ListItems.Add(, , "Clientes Relacionados")
    L.SubItems(1) = "013"
    
    Set L = LstReportes.ListItems.Add(, , "Declaracion Jurada Patrimonial")
    L.SubItems(1) = "014"
    
    Set L = LstReportes.ListItems.Add(, , "Documento REACTÍVATE PERÚ")
    L.SubItems(1) = "015"   'ANGC20200518
    
    ActxCta.NroCuenta = ""
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
End Sub

Private Sub LstCred_Click()
        If LstCred.ListCount > 0 And LstCred.ListIndex <> -1 Then
            ActxCta.NroCuenta = LstCred.Text
            ActxCta.SetFocusCuenta
        End If
End Sub

Public Sub ImprimePagareCred_X(ByVal psCtaCod As String)
Dim ssql As String
Dim oDCred As COMDCredito.DCOMCredito
Dim R As ADODB.Recordset
Dim RRelaCred As ADODB.Recordset
Dim sCadImp As String
Dim oFun As COMFunciones.FCOMCadenas
Dim nGaran As Integer
Dim nTitu As Integer
Dim nCode As Integer

Dim nTasaAnual As Double
Dim nTasaAnualMora As Double

Dim oWord As Word.Application
Dim oDoc As Word.Document
Dim oRange As Word.Range
Dim oApp As Excel.Application
Dim oLibro As Workbook
Dim oHoja As Worksheet
Dim oHoja2 As Worksheet
Dim nHoja2 As Boolean
Dim OCon As COMConecta.DCOMConecta
Dim sDistrito As String
Dim sSql2 As String
Dim R2 As ADODB.Recordset

    nHoja2 = False

    Set oDCred = New COMDCredito.DCOMCredito
    Set RRelaCred = oDCred.RecuperaRelacPers(psCtaCod)
    Set oDCred = Nothing
    
    Set oDCred = New COMDCredito.DCOMCredito
    Set R = oDCred.RecuperaDatosComunes(psCtaCod)
    Set oDCred = Nothing
    Set oFun = New COMFunciones.FCOMCadenas
  
    'Set oWord = CreateObject("Word.Application")
    '    oWord.Visible = True
    'Set oDoc = oWord.Documents.Open(App.path & "\FormatoCarta\FormatoPagare.doc")
    
    Set oApp = CreateObject("Excel.Application")
    Set oLibro = oApp.Workbooks.Open(App.Path & "\FormatoCarta\PAGARE.xlt")
    
    
    nTasaAnual = Round(((1 + (R!nTasaInteres / 100)) ^ 12 - 1) * 100, 2)
    nTasaAnualMora = Round(((1 + ((R!nTasaMora * 30) / 100)) ^ 12 - 1) * 100, 2)
    
    Set oHoja = oLibro.Sheets("PAGARE")
    With oHoja
    Do Until RRelaCred.EOF
    If RRelaCred!nOrden = 1 Then
        .Cells(3, 2) = psCtaCod
        If InStr(gsNomAge, "ESPINAR") > 0 Or InStr(gsNomAge, "AFLIGIDOS") > 0 Or InStr(gsNomAge, "WANCHAQ") > 0 Or InStr(gsNomAge, "SAN SEBASTIAN ") > 0 Then
            .Cells(3, 4) = "CUSCO"
        Else
            .Cells(3, 4) = Replace(gsNomAge, "AGENCIA", "")
        End If
        Set OCon = New COMConecta.DCOMConecta
        OCon.AbreConexion
        sSql2 = "Select cUbiGeoDescripcion From Ubicaciongeografica where cubigeocod like '3" & Mid(Trim(RRelaCred!cPersDireccUbiGeo), 2, Len(Trim(RRelaCred!cPersDireccUbiGeo)) - 2) & "%'"
        Set R2 = OCon.CargaRecordSet(sSql2)
        OCon.CierraConexion
        Set OCon = Nothing
        sDistrito = ""
        If R2.RecordCount > 0 Then
            sDistrito = R2!cUbiGeoDescripcion
        End If
        R2.Close
        Set R2 = Nothing
        
        
        '.Cells(3, 4) = gsNomAge
        .Cells(3, 8) = Day(gdFecSis)
        .Cells(3, 9) = Month(gdFecSis)
        .Cells(3, 10) = Year(gdFecSis)
        .Cells(3, 11) = Format(R!dVenc, "dd/mm/yyyy")
        .Cells(3, 11) = ""
        '.Cells(3, 13) = IIf(Mid(psCtaCod, 9, 1) = 1, "SOLES", "DOLARES") & "-" & Format(R!nMontoCol, "#0.00")
        .Cells(3, 13) = IIf(Mid(psCtaCod, 9, 1) = 1, " S/. ", "$/. ") & Format(R!nMontoCol, "#,0.00")
        .Cells(9, 4) = UCase(UnNumero(R!nMontoCol)) & " Y 00/100 " & IIf(Mid(psCtaCod, 9, 1) = 1, " NUEVOS SOLES", "DOLARES")
    
        .Cells(13, 3) = PstaNombre(RRelaCred!cPersNombre)
        .Cells(16, 3) = Trim(RRelaCred!cPersDireccDomicilio) & " - " & UCase(sDistrito)
        .Cells(19, 3) = IIf(IsNull(RRelaCred!DNI), RRelaCred!Ruc, RRelaCred!DNI)
        .Cells(19, 6) = IIf(IsNull(RRelaCred!cPersTelefono), "", RRelaCred!cPersTelefono)
        
    ElseIf RRelaCred!nOrden = 2 Then
        Set OCon = New COMConecta.DCOMConecta
        OCon.AbreConexion
        sSql2 = "Select cUbiGeoDescripcion From Ubicaciongeografica where cubigeocod like '3" & Mid(Trim(RRelaCred!cPersDireccUbiGeo), 2, Len(Trim(RRelaCred!cPersDireccUbiGeo)) - 2) & "%'"
        Set R2 = OCon.CargaRecordSet(sSql2)
        OCon.CierraConexion
        Set OCon = Nothing
        sDistrito = ""
        If R2.RecordCount > 0 Then
            sDistrito = R2!cUbiGeoDescripcion
        End If
        R2.Close
        Set R2 = Nothing

        .Cells(22, 3) = PstaNombre(RRelaCred!cPersNombre)
        .Cells(25, 3) = Trim(RRelaCred!cPersDireccDomicilio) & " - " & sDistrito
        .Cells(28, 3) = IIf(IsNull(RRelaCred!DNI), RRelaCred!Ruc, RRelaCred!DNI)
        .Cells(28, 6) = IIf(IsNull(RRelaCred!cPersTelefono), "", RRelaCred!cPersTelefono)
        
    ElseIf RRelaCred!nOrden = 3 Then
        nGaran = nGaran + 1
        If nGaran = 1 Then
            Set OCon = New COMConecta.DCOMConecta
            OCon.AbreConexion
            sSql2 = "Select cUbiGeoDescripcion From Ubicaciongeografica where cubigeocod like '3" & Mid(Trim(RRelaCred!cPersDireccUbiGeo), 2, Len(Trim(RRelaCred!cPersDireccUbiGeo)) - 3) & "%'"
            Set R2 = OCon.CargaRecordSet(sSql2)
            OCon.CierraConexion
            Set OCon = Nothing
            sDistrito = ""
            If R2.RecordCount > 0 Then
                sDistrito = R2!cUbiGeoDescripcion
            End If
            R2.Close
            Set R2 = Nothing
            
            .Cells(31, 3) = PstaNombre(RRelaCred!cPersNombre)
            .Cells(33, 4) = ""
            .Cells(35, 3) = Trim(RRelaCred!cPersDireccDomicilio) & " - " & sDistrito
            .Cells(37, 3) = IIf(IsNull(RRelaCred!DNI), RRelaCred!Ruc, RRelaCred!DNI)
            .Cells(37, 6) = IIf(IsNull(RRelaCred!cPersTelefono), "", RRelaCred!cPersTelefono)
        ElseIf nGaran = 2 Then
            Set OCon = New COMConecta.DCOMConecta
            OCon.AbreConexion
            sSql2 = "Select cUbiGeoDescripcion From Ubicaciongeografica where cubigeocod like '3" & Mid(Trim(RRelaCred!cPersDireccUbiGeo), 2, Len(Trim(RRelaCred!cPersDireccUbiGeo)) - 3) & "%'"
            Set R2 = OCon.CargaRecordSet(sSql2)
            OCon.CierraConexion
            Set OCon = Nothing
            sDistrito = ""
            If R2.RecordCount > 0 Then
                sDistrito = R2!cUbiGeoDescripcion
            End If
            R2.Close
            Set R2 = Nothing
            
            .Cells(39, 3) = PstaNombre(RRelaCred!cPersNombre)
            .Cells(41, 4) = ""
            .Cells(43, 3) = Trim(RRelaCred!cPersDireccDomicilio) & " - " & sDistrito
            .Cells(45, 3) = IIf(IsNull(RRelaCred!DNI), RRelaCred!Ruc, RRelaCred!DNI)
            .Cells(45, 6) = IIf(IsNull(RRelaCred!cPersTelefono), "", RRelaCred!cPersTelefono)
            
        ElseIf nGaran = 3 Then
            Set oHoja2 = oLibro.Sheets("PAGARE2")
            nHoja2 = True
            With oHoja2
            
            Set OCon = New COMConecta.DCOMConecta
            OCon.AbreConexion
            sSql2 = "Select cUbiGeoDescripcion From Ubicaciongeografica where cubigeocod like '3" & Mid(Trim(RRelaCred!cPersDireccUbiGeo), 2, Len(Trim(RRelaCred!cPersDireccUbiGeo)) - 3) & "%'"
            Set R2 = OCon.CargaRecordSet(sSql2)
            OCon.CierraConexion
            Set OCon = Nothing
            sDistrito = ""
            If R2.RecordCount > 0 Then
                sDistrito = R2!cUbiGeoDescripcion
            End If
            R2.Close
            Set R2 = Nothing
                    
            
                .Cells(3, 1) = PstaNombre(RRelaCred!cPersNombre)
                .Cells(3, 2) = ""
                .Cells(6, 1) = Trim(RRelaCred!cPersDireccDomicilio) & " - " & sDistrito
                .Cells(8, 1) = IIf(IsNull(RRelaCred!DNI), RRelaCred!Ruc, RRelaCred!DNI)
                .Cells(8, 4) = IIf(IsNull(RRelaCred!cPersTelefono), "", RRelaCred!cPersTelefono)
                
            End With
        ElseIf nGaran = 4 Then
            
            Set OCon = New COMConecta.DCOMConecta
            OCon.AbreConexion
            sSql2 = "Select cUbiGeoDescripcion From Ubicaciongeografica where cubigeocod like '3" & Mid(Trim(RRelaCred!cPersDireccUbiGeo), 2, Len(Trim(RRelaCred!cPersDireccUbiGeo)) - 3) & "%'"
            Set R2 = OCon.CargaRecordSet(sSql2)
            OCon.CierraConexion
            Set OCon = Nothing
            sDistrito = ""
            If R2.RecordCount > 0 Then
                sDistrito = R2!cUbiGeoDescripcion
            End If
            R2.Close
            Set R2 = Nothing
        
            Set oHoja2 = oLibro.Sheets("PAGARE2")
            nHoja2 = True
            With oHoja2
                .Cells(3, 6) = PstaNombre(RRelaCred!cPersNombre)
                .Cells(3, 2) = ""
                .Cells(6, 6) = Trim(RRelaCred!cPersDireccDomicilio) & " - " & sDistrito
                .Cells(8, 6) = IIf(IsNull(RRelaCred!DNI), RRelaCred!Ruc, RRelaCred!DNI)
                .Cells(8, 9) = IIf(IsNull(RRelaCred!cPersTelefono), "", RRelaCred!cPersTelefono)
                
            End With
        End If
    End If
    RRelaCred.MoveNext
    Loop
        
        .Cells(29, 13) = Format(nTasaAnual, "#0.00") & "%"
        .Cells(39, 11) = Format(nTasaAnualMora, "#0.00") & "%"
        
    End With
    
    Dim P As Printer
    Dim Impresora As String
    
    Impresora = ""
'    For Each P In Printers
'        If InStr(1, UCase(P.DeviceName), "HP") <> 0 Then
'            Impresora = P.DeviceName
'            Exit For
'        End If
'    Next
    Impresora = sLpt
    'oApp.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, ActivePrinter:=Impresora
    'oApp.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, ActivePrinter:=Impresora
    
    oHoja.SaveAs App.Path & "\SPOOLER\" & psCtaCod
    'oHoja2.SaveAs App.path & "\SPOOLER\" & psCtaCod & "_2"
    
    oApp.Visible = True
    oApp.Windows(1).Visible = True
    
    'oHoja.PrintOut Copies:=1, Collate:=True, ActivePrinter:=Impresora
    'If nHoja2 = True Then
    '    oHoja2.PrintOut Copies:=1, Collate:=True, ActivePrinter:=Impresora
    'End If
        
    'olibro.Close False
    
    Set oHoja = Nothing
    Set oLibro = Nothing
    'oApp.Quit
    Set oApp = Nothing
    
'
'    With oWord.Selection.Find
'        .Text = "cNroCta"
'        .Replacement.Text = psCtaCod
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .Execute Replace:=wdReplaceAll
'    End With
'
'    With oWord.Selection.Find
'        .Text = "cLugarE"
'        .Replacement.Text = gsNomAge
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .Execute Replace:=wdReplaceAll
'    End With
'
'    With oWord.Selection.Find
'        .Text = "cDia"
'        .Replacement.Text = Day(gdFecSis)
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .Execute Replace:=wdReplaceAll
'    End With
'
'    With oWord.Selection.Find
'        .Text = "cMes"
'        .Replacement.Text = Month(gdFecSis)
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .Execute Replace:=wdReplaceAll
'    End With
'
'    With oWord.Selection.Find
'        .Text = "cAño"
'        .Replacement.Text = Year(gdFecSis)
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .Execute Replace:=wdReplaceAll
'    End With
'
'    With oWord.Selection.Find
'        .Text = "cTasa"
'        .Replacement.Text = Format(nTasaAnual, "0.0")
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .Execute Replace:=wdReplaceAll
'    End With
'
'    If Not (R.EOF And R.BOF) Then
'        With oWord.Selection.Find
'            .Text = "cFecVen"
'            .Replacement.Text = R!dVenc
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'            .Text = "cMonto"
'            .Replacement.Text = IIf(Mid(psCtaCod, 9, 1) = 1, "SOLES", "DOLARES") & "-" & R!nMontoCol
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'     End If
'
'     With oWord.Selection.Find
'        .Text = "cMLetra"
'        .Replacement.Text = UnNumero(R!nMontoCol)
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .Execute Replace:=wdReplaceAll
'     End With
'
'    nGaran = 0
'
'
'    If Not (RRelaCred.EOF And RRelaCred.BOF) Then
'       Do Until RRelaCred.EOF
'          If RRelaCred!nOrden = 1 Then
'               With oWord.Selection.Find
'                    .Text = "cNomTitu"
'                    .Replacement.Text = PstaNombre(RRelaCred!cPersNombre)
'                    .Forward = True
'                    .Wrap = wdFindContinue
'                    .Format = False
'                    .Execute Replace:=wdReplaceAll
'               End With
'
'               With oWord.Selection.Find
'                    .Text = "cDNITitu"
'                    .Replacement.Text = IIf(IsNull(RRelaCred!DNI), RRelaCred!RUC, RRelaCred!DNI)
'                    .Forward = True
'                    .Wrap = wdFindContinue
'                    .Format = False
'                    .Execute Replace:=wdReplaceAll
'               End With
'
'               With oWord.Selection.Find
'                    .Text = "cTelefonoTitu"
'                    .Replacement.Text = IIf(IsNull(RRelaCred!cPersTelefono), "", RRelaCred!cPersTelefono)
'                    .Forward = True
'                    .Wrap = wdFindContinue
'                    .Format = False
'                    .Execute Replace:=wdReplaceAll
'               End With
'
'
'               With oWord.Selection.Find
'                    .Text = "cDirTitu"
'                    .Replacement.Text = RRelaCred!cPersDireccDomicilio
'                    .Forward = True
'                    .Wrap = wdFindContinue
'                    .Format = False
'                    .Execute Replace:=wdReplaceAll
'               End With
'               nTitu = 1
'          ElseIf RRelaCred!nOrden = 2 Then
'               With oWord.Selection.Find
'                    .Text = "cNomCode"
'                    .Replacement.Text = PstaNombre(RRelaCred!cPersNombre)
'                    .Forward = True
'                    .Wrap = wdFindContinue
'                    .Format = False
'                    .Execute Replace:=wdReplaceAll
'               End With
'
'               With oWord.Selection.Find
'                    .Text = "cDNICode"
'                    .Replacement.Text = IIf(IsNull(RRelaCred!DNI), RRelaCred!RUC, RRelaCred!DNI)
'                    .Forward = True
'                    .Wrap = wdFindContinue
'                    .Format = False
'                    .Execute Replace:=wdReplaceAll
'               End With
'
'               With oWord.Selection.Find
'                    .Text = "cTelefonoCode"
'                    .Replacement.Text = IIf(IsNull(RRelaCred!cPersTelefono), "", RRelaCred!cPersTelefono)
'                    .Forward = True
'                    .Wrap = wdFindContinue
'                    .Format = False
'                    .Execute Replace:=wdReplaceAll
'               End With
'
'
'               With oWord.Selection.Find
'                    .Text = "cDirCode"
'                    .Replacement.Text = RRelaCred!cPersDireccDomicilio
'                    .Forward = True
'                    .Wrap = wdFindContinue
'                    .Format = False
'                    .Execute Replace:=wdReplaceAll
'               End With
'               nCode = 1
'          ElseIf RRelaCred!nOrden = 3 Then
'              nGaran = nGaran + 1
'              If nGaran = 1 Then
'                  With oWord.Selection.Find
'                       .Text = "cNomGara1"
'                       .Replacement.Text = PstaNombre(RRelaCred!cPersNombre)
'                       .Forward = True
'                       .Wrap = wdFindContinue
'                       .Format = False
'                       .Execute Replace:=wdReplaceAll
'                  End With
'
'                  With oWord.Selection.Find
'                       .Text = "cDNIGara1"
'                       .Replacement.Text = IIf(IsNull(RRelaCred!DNI), RRelaCred!RUC, RRelaCred!DNI)
'                       .Forward = True
'                       .Wrap = wdFindContinue
'                       .Format = False
'                       .Execute Replace:=wdReplaceAll
'                  End With
'
'                  With oWord.Selection.Find
'                       .Text = "cTelefonoGara1"
'                       .Replacement.Text = IIf(IsNull(RRelaCred!cPersTelefono), "", RRelaCred!cPersTelefono)
'                       .Forward = True
'                       .Wrap = wdFindContinue
'                       .Format = False
'                       .Execute Replace:=wdReplaceAll
'                  End With
'
'
'                  With oWord.Selection.Find
'                       .Text = "cDirGara1"
'                       .Replacement.Text = RRelaCred!cPersDireccDomicilio
'                       .Forward = True
'                       .Wrap = wdFindContinue
'                       .Format = False
'                       .Execute Replace:=wdReplaceAll
'                  End With
'              ElseIf nGaran = 2 Then
'                  With oWord.Selection.Find
'                       .Text = "cNomGara2"
'                       .Replacement.Text = PstaNombre(RRelaCred!cPersNombre)
'                       .Forward = True
'                       .Wrap = wdFindContinue
'                       .Format = False
'                       .Execute Replace:=wdReplaceAll
'                  End With
'
'                  With oWord.Selection.Find
'                       .Text = "cDNIGara2"
'                       .Replacement.Text = IIf(IsNull(RRelaCred!DNI), RRelaCred!RUC, RRelaCred!DNI)
'                       .Forward = True
'                       .Wrap = wdFindContinue
'                       .Format = False
'                       .Execute Replace:=wdReplaceAll
'                  End With
'
'                  With oWord.Selection.Find
'                       .Text = "cTelefonoGara2"
'                       .Replacement.Text = IIf(IsNull(RRelaCred!cPersTelefono), "", RRelaCred!cPersTelefono)
'                       .Forward = True
'                       .Wrap = wdFindContinue
'                       .Format = False
'                       .Execute Replace:=wdReplaceAll
'                  End With
'
'
'                  With oWord.Selection.Find
'                       .Text = "cDirGara2"
'                       .Replacement.Text = RRelaCred!cPersDireccDomicilio
'                       .Forward = True
'                       .Wrap = wdFindContinue
'                       .Format = False
'                       .Execute Replace:=wdReplaceAll
'                  End With
'
'              ElseIf nGaran = 3 Then
'                  With oWord.Selection.Find
'                       .Text = "cNomGara3"
'                       .Replacement.Text = PstaNombre(RRelaCred!cPersNombre)
'                       .Forward = True
'                       .Wrap = wdFindContinue
'                       .Format = False
'                       .Execute Replace:=wdReplaceAll
'                  End With
'
'                  With oWord.Selection.Find
'                       .Text = "cDNIGara3"
'                       .Replacement.Text = IIf(IsNull(RRelaCred!DNI), RRelaCred!RUC, RRelaCred!DNI)
'                       .Forward = True
'                       .Wrap = wdFindContinue
'                       .Format = False
'                       .Execute Replace:=wdReplaceAll
'                  End With
'
'                  With oWord.Selection.Find
'                       .Text = "cTelefonoGara3"
'                       .Replacement.Text = IIf(IsNull(RRelaCred!cPersTelefono), "", RRelaCred!cPersTelefono)
'                       .Forward = True
'                       .Wrap = wdFindContinue
'                       .Format = False
'                       .Execute Replace:=wdReplaceAll
'                  End With
'
'
'                  With oWord.Selection.Find
'                       .Text = "cDirGara3"
'                       .Replacement.Text = RRelaCred!cPersDireccDomicilio
'                       .Forward = True
'                       .Wrap = wdFindContinue
'                       .Format = False
'                       .Execute Replace:=wdReplaceAll
'                  End With
'              ElseIf nGaran = 4 Then
'                  With oWord.Selection.Find
'                       .Text = "cNomGara4"
'                       .Replacement.Text = PstaNombre(RRelaCred!cPersNombre)
'                       .Forward = True
'                       .Wrap = wdFindContinue
'                       .Format = False
'                       .Execute Replace:=wdReplaceAll
'                  End With
'
'                  With oWord.Selection.Find
'                       .Text = "cDNIGara4"
'                       .Replacement.Text = IIf(IsNull(RRelaCred!DNI), RRelaCred!RUC, RRelaCred!DNI)
'                       .Forward = True
'                       .Wrap = wdFindContinue
'                       .Format = False
'                       .Execute Replace:=wdReplaceAll
'                  End With
'
'                  With oWord.Selection.Find
'                       .Text = "cTelefonoGara4"
'                       .Replacement.Text = IIf(IsNull(RRelaCred!cPersTelefono), "", RRelaCred!cPersTelefono)
'                       .Forward = True
'                       .Wrap = wdFindContinue
'                       .Format = False
'                       .Execute Replace:=wdReplaceAll
'                  End With
'
'
'                  With oWord.Selection.Find
'                       .Text = "cDirGara4"
'                       .Replacement.Text = RRelaCred!cPersDireccDomicilio
'                       .Forward = True
'                       .Wrap = wdFindContinue
'                       .Format = False
'                       .Execute Replace:=wdReplaceAll
'                  End With
'              ElseIf nGaran = 5 Then
'                  With oWord.Selection.Find
'                       .Text = "cNomGara5"
'                       .Replacement.Text = PstaNombre(RRelaCred!cPersNombre)
'                       .Forward = True
'                       .Wrap = wdFindContinue
'                       .Format = False
'                       .Execute Replace:=wdReplaceAll
'                  End With
'
'                  With oWord.Selection.Find
'                       .Text = "cDNIGara5"
'                       .Replacement.Text = IIf(IsNull(RRelaCred!DNI), RRelaCred!RUC, RRelaCred!DNI)
'                       .Forward = True
'                       .Wrap = wdFindContinue
'                       .Format = False
'                       .Execute Replace:=wdReplaceAll
'                  End With
'
'                  With oWord.Selection.Find
'                       .Text = "cTelefonoGara5"
'                       .Replacement.Text = IIf(IsNull(RRelaCred!cPersTelefono), "", RRelaCred!cPersTelefono)
'                       .Forward = True
'                       .Wrap = wdFindContinue
'                       .Format = False
'                       .Execute Replace:=wdReplaceAll
'                  End With
'
'
'                  With oWord.Selection.Find
'                       .Text = "cDirGara5"
'                       .Replacement.Text = RRelaCred!cPersDireccDomicilio
'                       .Forward = True
'                       .Wrap = wdFindContinue
'                       .Format = False
'                       .Execute Replace:=wdReplaceAll
'                  End With
'              ElseIf nGaran = 6 Then
'                  With oWord.Selection.Find
'                       .Text = "cNomGara6"
'                       .Replacement.Text = PstaNombre(RRelaCred!cPersNombre)
'                       .Forward = True
'                       .Wrap = wdFindContinue
'                       .Format = False
'                       .Execute Replace:=wdReplaceAll
'                  End With
'
'                  With oWord.Selection.Find
'                       .Text = "cDNIGara6"
'                       .Replacement.Text = IIf(IsNull(RRelaCred!DNI), RRelaCred!RUC, RRelaCred!DNI)
'                       .Forward = True
'                       .Wrap = wdFindContinue
'                       .Format = False
'                       .Execute Replace:=wdReplaceAll
'                  End With
'
'                  With oWord.Selection.Find
'                       .Text = "cTelefonoGara6"
'                       .Replacement.Text = IIf(IsNull(RRelaCred!cPersTelefono), "", RRelaCred!cPersTelefono)
'                       .Forward = True
'                       .Wrap = wdFindContinue
'                       .Format = False
'                       .Execute Replace:=wdReplaceAll
'                  End With
'
'
'                  With oWord.Selection.Find
'                       .Text = "cDirGara6"
'                       .Replacement.Text = RRelaCred!cPersDireccDomicilio
'                       .Forward = True
'                       .Wrap = wdFindContinue
'                       .Format = False
'                       .Execute Replace:=wdReplaceAll
'                  End With
'              End If
'
'          End If
'          RRelaCred.MoveNext
'       Loop
'    End If
'    If nTitu = 0 Then
'        With oWord.Selection.Find
'            .Text = "cNomTitu"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cDNITitu"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cTelefonoTitu"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'            .Text = "cDirTitu"
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'    End If
'
'    If nCode = 0 Then
'        With oWord.Selection.Find
'            .Text = "cNomCode"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cDNICode"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cTelefonoCode"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'            .Text = "cDirCode"
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'    End If
'
'    If nGaran = 0 Then
'        With oWord.Selection.Find
'             .Text = "cNomGara1"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cDNIGara1"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cTelefonoGara1"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'            .Text = "cDirGara1"
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cNomGara2"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cDNIGara2"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cTelefonoGara2"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'            .Text = "cDirGara2"
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cNomGara3"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cDNIGara3"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cTelefonoGara3"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'            .Text = "cDirGara3"
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cNomGara4"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cDNIGara4"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cTelefonoGara4"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'            .Text = "cDirGara4"
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cNomGara5"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cDNIGara5"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cTelefonoGara5"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'            .Text = "cDirGara5"
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cNomGara6"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cDNIGara6"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cTelefonoGara6"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'            .Text = "cDirGara6"
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'    ElseIf nGaran = 2 Then
'        With oWord.Selection.Find
'             .Text = "cNomGara3"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cDNIGara3"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cTelefonoGara3"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'            .Text = "cDirGara3"
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cNomGara4"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cDNIGara4"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cTelefonoGara4"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'            .Text = "cDirGara4"
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cNomGara5"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cDNIGara5"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cTelefonoGara5"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'            .Text = "cDirGara5"
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cNomGara6"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cDNIGara6"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cTelefonoGara6"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'            .Text = "cDirGara6"
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'    ElseIf nGaran = 1 Then
'        With oWord.Selection.Find
'             .Text = "cNomGara2"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cDNIGara2"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cTelefonoGara2"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'            .Text = "cDirGara2"
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cNomGara3"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cDNIGara3"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cTelefonoGara3"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'            .Text = "cDirGara3"
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cNomGara4"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cDNIGara4"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cTelefonoGara4"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'            .Text = "cDirGara4"
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cNomGara5"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cDNIGara5"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cTelefonoGara5"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'            .Text = "cDirGara5"
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cNomGara6"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cDNIGara6"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cTelefonoGara6"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'            .Text = "cDirGara6"
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'    ElseIf nGaran = 3 Then
'        With oWord.Selection.Find
'             .Text = "cNomGara4"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cDNIGara4"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cTelefonoGara4"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'            .Text = "cDirGara4"
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cNomGara5"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cDNIGara5"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cTelefonoGara5"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'            .Text = "cDirGara5"
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cNomGara6"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cDNIGara6"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cTelefonoGara6"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'            .Text = "cDirGara6"
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'    ElseIf nGaran = 4 Then
'        With oWord.Selection.Find
'             .Text = "cNomGara5"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cDNIGara5"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cTelefonoGara5"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'            .Text = "cDirGara5"
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cNomGara6"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cDNIGara6"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cTelefonoGara6"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'            .Text = "cDirGara6"
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'    ElseIf nGaran = 5 Then
'        With oWord.Selection.Find
'             .Text = "cNomGara6"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cDNIGara6"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'             .Text = "cTelefonoGara6"
'             .Replacement.Text = ""
'             .Forward = True
'             .Wrap = wdFindContinue
'             .Format = False
'             .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'            .Text = "cDirGara6"
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'    End If
'
    
    
End Sub

'*** PEAC 20080623
Public Function ImprimeDDJJPatrimonial(ByVal pRsAgencia As ADODB.Recordset, ByVal pRsDatosDDJJ As ADODB.Recordset, _
                    ByVal pRsGarantTitular As ADODB.Recordset, ByVal pRsGarantAvales As ADODB.Recordset, ByVal pdFecSis As Date) As String
  
Dim lcPersCod As String, lcDNI As String, lcDniconyuge As String, lcPersNombre As String, lcConyuge As String
Dim lnTotSoles As Double, lnTotDolar As Double, lcNumGarant As String, lnTotalDJ As Double
Dim lcNumGarantA As String

Dim Cuenta As Integer

   Dim Word As New Word.Application
    
    'AGREGA  DOCUMENTO
    Word.Documents.Add
    
    Dim X As Integer

'    Word.Selection.PageSetup.LeftMargin = CentimetersToPoints(1.5)
'    Word.Selection.PageSetup.RightMargin = CentimetersToPoints(1)

    Word.Selection.Font.Name = "Courier New"
    Word.Selection.Font.Size = 9

'   Word.Selection.Font.Color = wdColorRed
'   Word.Selection.TypeText "Párrafo " & x & vbCrLf
    
    Word.Selection.Font.Bold = wdToggle
    Word.Selection.TypeText "CAJA MAYNAS"
    Word.Selection.TypeParagraph
    Word.Selection.TypeText "RUC: 20103845328"
    Word.Selection.TypeParagraph
    Word.Selection.TypeParagraph
    Word.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Word.Selection.TypeText "DECLARACION JURADA PATRIMONIAL"
    Word.Selection.TypeParagraph
    Word.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    Word.Selection.Font.Bold = wdToggle
    Word.Selection.TypeParagraph
    Word.Selection.TypeParagraph

    
    If Len(pRsDatosDDJJ!DNI) > 0 Then
        Word.Selection.TypeText "YO, " & Trim(pRsDatosDDJJ!cPersNombre) & ", IDENTIFICADO CON DNI No " & Trim(pRsDatosDDJJ!DNI) & " DE ESTADO CIVIL, " & IIf(Len(pRsDatosDDJJ!EstadoCiv) > 0, Trim(pRsDatosDDJJ!EstadoCiv), "**********") & _
         " Y (nombre del conyugue, si fuese el caso) " & IIf(Len(pRsDatosDDJJ!Conyuge) > 0, Trim(pRsDatosDDJJ!Conyuge), IIf(Len(pRsDatosDDJJ!Codeudor) > 0, Trim(pRsDatosDDJJ!Codeudor), "**********")) & " CON DNI No " & IIf(Len(pRsDatosDDJJ!dniconyuge) > 0, Trim(pRsDatosDDJJ!dniconyuge), IIf(Len(pRsDatosDDJJ!dnicodeudor) > 0, Trim(pRsDatosDDJJ!dnicodeudor), "**********")) & _
         " DECLARO (AMOS) BAJO JURAMENTO DE LEALTAD Y HONESTIDAD QUE SOY (SOMOS) PROPIETARIO(S) DE LOS " & _
         " SIGUIENTES BIENES:"
    Else
        Word.Selection.TypeText "LA EMPRESA, " & Trim(pRsDatosDDJJ!cPersNombre) & ", CON RUC No " & Trim(pRsDatosDDJJ!Ruc) & _
            " DECLARA BAJO JURAMENTO DE LEALTAD Y HONESTIDAD QUE SOY (SOMOS) PROPIETARIO(S) DE LOS " & _
            " SIGUIENTES BIENES:"
    End If
    Word.Selection.TypeParagraph
    Word.Selection.TypeParagraph
    
    Word.Selection.TypeText "DESCRIPCION                     DIRECCION                    VALOR REALIZACION  MONEDA"
    Word.Selection.TypeParagraph
    Word.Selection.TypeText "======================================================================================"
    Word.Selection.TypeParagraph

    lnTotSoles = 0
    lnTotDolar = 0
    
    lnTotalDJ = 0
    
    Do While Not pRsGarantTitular.EOF
    
        lcNumGarant = pRsGarantTitular!cNumGarant
        lnTotalDJ = pRsGarantTitular!TotalDJ
    
        Word.Selection.TypeText Replace(Left(UCase(Replace(Trim(pRsGarantTitular!cDescripcion), "#º", " ")) & Space(30), 30), " ", " ") & Space(2) & Replace(Left(Replace(Replace(UCase(Trim(pRsGarantTitular!cDireccion)), "#", "No"), "º", "o") & Space(30), 30), " ", " ") & Space(2) & _
            IIf(pRsGarantTitular!nTpoGarantia = 4, Space(14), Right(Space(14) & Format(pRsGarantTitular!nRealizacion, "#,#0.00"), 10)) & Space(2) & _
            IIf(pRsGarantTitular!nTpoGarantia = 4, Space(5), Replace(Left(UCase(pRsGarantTitular!Moneda) & Space(5), 5), " ", " "))
        Word.Selection.TypeParagraph
        
        If pRsGarantTitular!nTpoGarantia <> 4 Then
            Word.Selection.TypeParagraph
        End If
        
        '*********SI ES ARTEFACTO AUMENTAR UNA FILA

        If pRsGarantTitular!nmoneda = 1 Then
            If pRsGarantTitular!nTpoGarantia <> 4 Then
                lnTotSoles = lnTotSoles + pRsGarantTitular!nRealizacion
            End If
        Else
            If pRsGarantTitular!nTpoGarantia <> 4 Then
                lnTotDolar = lnTotDolar + pRsGarantTitular!nRealizacion
            End If
        End If
        
        Do While pRsGarantTitular!cNumGarant = lcNumGarant And lnTotalDJ > 0
            
            Word.Selection.TypeText Space(6) & Right(Space(5) & Format(pRsGarantTitular!nItem, "#,#0"), 5) & Space(2) & Replace(Left(UCase(Trim(pRsGarantTitular!cGarDjDescripcion)) & Space(30), 30), " ", " ") & Space(2) & Right(Space(5) & Format(pRsGarantTitular!nGarDJCantidad, "#,#0"), 5) & Space(12) & Right(Space(12) & Format(pRsGarantTitular!TotalDJ, "#,#0.00"), 12) & Space(2) & Replace(Left(UCase(pRsGarantTitular!Moneda) & Space(5), 5), " ", " ")
            Word.Selection.TypeParagraph
        
            If pRsGarantTitular!nmoneda = 1 Then
                lnTotSoles = lnTotSoles + pRsGarantTitular!TotalDJ
            Else
                lnTotDolar = lnTotDolar + pRsGarantTitular!TotalDJ
            End If
        
        
            pRsGarantTitular.MoveNext
            If pRsGarantTitular.EOF Then
                Exit Do
            End If
        Loop
        
        Word.Selection.TypeParagraph
        
        If pRsGarantTitular.EOF Then
            Exit Do
        Else
            If pRsGarantTitular!cNumGarant = lcNumGarant Then
                pRsGarantTitular.MoveNext
            Else
                Word.Selection.TypeParagraph
            End If
        End If
    
    Loop
    pRsGarantTitular.Close
    Set pRsGarantTitular = Nothing
    
    'Word.Selection.TypeParagraph
    
    Word.Selection.TypeText "======================================================================================"
    Word.Selection.TypeParagraph
    Word.Selection.TypeText "TOTAL PATRIMONIO DECLARADO - SOLES :" & Right(Space(14) & Format(lnTotSoles, "#,#0.00"), 14)
    Word.Selection.TypeParagraph
    Word.Selection.TypeText "TOTAL PATRIMONIO DECLARADO - DOLAR :" & Right(Space(14) & Format(lnTotDolar, "#,#0.00"), 14)
    Word.Selection.TypeParagraph
    Word.Selection.TypeParagraph
    
    Word.Selection.TypeText "DEJO (AMOS) PLENA CONSTANCIA DE LOS BIENES DECLARADOS LO (LOS) OFREZCO (CEMOS) IRREVOCABLEMENTE  EN RESPALDO A  LOS" & _
                " CREDITOS DIRECTOS E INDIRECTOS QUE TENGA A BIEN CONCEDERME (NOS) LA CAJA MUNICIPAL DE AHORRO Y CREDITO DE  MAYNAS." & _
                " DECLARO (MOS) QUE LA INFORMACION PROPORCIONADA EN ESTE DOCUMENTO ES AUTENTICA Y AUTORIZO (AMOS) A LA CAJA MUNICIPAL" & _
                " A REALIZAR SU VERIFICACION. ESTE DOCUMENTO TIENE CARACTER DE DECLARACION JURADA DE ACUERDO CON EL ARTICULO 179 DE" & _
                " LA LEY No 26702 DECLARO ASI MISMO QUE EN CASO DE PROPORCIONAR INFORMACION FALSA ESTARE COMPRENDIDO DENTRO  DE LOS" & _
                " ALCANCES  DEL  ART. 247 DEL  CODIGO PENAL QUE  ESTABLECE LA TIPIFICACION  DEL DELITO  FINANCIERO POR PROPORCIONAR" & _
                " INFORMACION FALSA A UNA EMPRESA DEL SISTEMA FINANCIERO."
    Word.Selection.TypeParagraph
    Word.Selection.TypeParagraph
    
    Word.Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
    Word.Selection.TypeText Trim(pRsAgencia!Dist) & ", " & Format(pdFecSis, "dddd") & ", " & Day(pdFecSis) & " de " & Format(pdFecSis, "mmmm") & " de " & Year(pdFecSis)
    Word.Selection.TypeParagraph
    Word.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    Word.Selection.TypeParagraph
    Word.Selection.TypeParagraph
    Word.Selection.TypeParagraph
    Word.Selection.TypeParagraph

    If Len(pRsDatosDDJJ!DNI) > 0 Then
        If Len(pRsDatosDDJJ!dniconyuge) > 0 Then
            Word.Selection.TypeText ".............................." & Space(22) & ".............................."
            Word.Selection.TypeParagraph
            Word.Selection.TypeText "DECLARANTE:" & Trim(pRsDatosDDJJ!cPersNombre) & Space(2) & "CONYUGE: " & Trim(pRsDatosDDJJ!Conyuge)
            Word.Selection.TypeParagraph
            Word.Selection.TypeText "DNI:" & Trim(pRsDatosDDJJ!DNI) & Space(44) & "DNI: " & Trim(pRsDatosDDJJ!dniconyuge)
            Word.Selection.TypeParagraph
            
        ElseIf Len(pRsDatosDDJJ!dnicodeudor) > 0 Then
            Word.Selection.TypeText ".............................." & Space(22) & ".............................."
            Word.Selection.TypeParagraph
            Word.Selection.TypeText "DECLARANTE:" & Trim(pRsDatosDDJJ!cPersNombre) & Space(2) & "CONYUGE: " & Trim(pRsDatosDDJJ!Codeudor)
            Word.Selection.TypeParagraph
            Word.Selection.TypeText "DNI:" & Trim(pRsDatosDDJJ!DNI) & Space(44) & "DNI: " & Trim(pRsDatosDDJJ!dnicodeudor)
            Word.Selection.TypeParagraph
        Else
            Word.Selection.TypeText ".............................."
            Word.Selection.TypeParagraph
            Word.Selection.TypeText "DECLARANTE:" & Trim(pRsDatosDDJJ!cPersNombre)
            Word.Selection.TypeParagraph
            Word.Selection.TypeText "DNI:" & Trim(pRsDatosDDJJ!DNI)
            Word.Selection.TypeParagraph
        End If
    Else
        Word.Selection.TypeText ".............................."
        Word.Selection.TypeParagraph
        Word.Selection.TypeText "DECLARANTE:" & Trim(pRsDatosDDJJ!cPersNombre)
        Word.Selection.TypeParagraph
        Word.Selection.TypeText "RUC: " & Trim(pRsDatosDDJJ!Ruc)
        Word.Selection.TypeParagraph
    End If

'*************************---------------------- Imprime las garantias de los avales

'*** PEAC 20080730

If Not (pRsGarantAvales.BOF And pRsGarantAvales.EOF) Then

    Do While Not pRsGarantAvales.EOF
    
        Word.Selection.InsertBreak Type:=wdPageBreak
        
        lcPersCod = pRsGarantAvales!cperscod
        lcDNI = pRsGarantAvales!DNI
        lcDniconyuge = pRsGarantAvales!dniconyuge
        lcPersNombre = pRsGarantAvales!cPersNombre
        lcConyuge = pRsGarantAvales!Conyuge
        
        Word.Selection.Font.Bold = wdToggle
        Word.Selection.TypeText "CAJA MAYNAS"
        Word.Selection.TypeParagraph
        Word.Selection.TypeText "RUC: 20103845328"
        Word.Selection.TypeParagraph
        Word.Selection.TypeParagraph
        Word.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        Word.Selection.TypeText "DECLARACION JURADA PATRIMONIAL"
        Word.Selection.TypeParagraph
        Word.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
        Word.Selection.Font.Bold = wdToggle
        Word.Selection.TypeParagraph
        Word.Selection.TypeParagraph


        Word.Selection.TypeText "YO, " & Trim(pRsGarantAvales!cPersNombre) & ", IDENTIFICADO CON DNI No " & Trim(pRsGarantAvales!DNI) & " DE ESTADO CIVIL, " & IIf(Len(pRsGarantAvales!EstadoCiv) > 0, Trim(pRsGarantAvales!EstadoCiv), "**********") & _
                    " Y (nombre del conyugue, si fuese el caso) " & IIf(Len(pRsGarantAvales!Conyuge) > 0, Trim(pRsGarantAvales!Conyuge), "**********") & " CON DNI No " & IIf(Len(pRsGarantAvales!dniconyuge) > 0, Trim(pRsGarantAvales!dniconyuge), "**********") & _
                    " DECLARO (AMOS) BAJO JURAMENTO DE LEALTAD Y HONESTIDAD QUE SOY (SOMOS) PROPIETARIO(S) DE LOS " & _
                    " SIGUIENTES BIENES:"
        Word.Selection.TypeParagraph
        Word.Selection.TypeParagraph
        
        Word.Selection.TypeText "DESCRIPCION                     DIRECCION                    VALOR REALIZACION  MONEDA"
        Word.Selection.TypeParagraph
        Word.Selection.TypeText "======================================================================================"
        Word.Selection.TypeParagraph
        
        lnTotSoles = 0
        lnTotDolar = 0
        lnTotalDJ = 0
        
        Do While pRsGarantAvales!cperscod = lcPersCod
            lcNumGarantA = pRsGarantAvales!cNumGarant
            lnTotalDJ = pRsGarantAvales!TotalDJ
            
            Word.Selection.TypeText Replace(Left(UCase(Trim(pRsGarantAvales!cDescripcion)) & Space(30), 30), " ", " ") & Space(2) & Replace(Left(UCase(Trim(pRsGarantAvales!cDireccion)) & Space(30), 30), " ", " ") & Space(2) & _
            IIf(pRsGarantAvales!nTpoGarantia = 4, Space(14), Right(Space(14) & Format(pRsGarantAvales!nRealizacion, "#,#0.00"), 10)) & Space(2) & _
            IIf(pRsGarantAvales!nTpoGarantia = 4, Space(5), Replace(Left(UCase(pRsGarantAvales!Moneda) & Space(5), 5), " ", " "))

            Word.Selection.TypeParagraph
                        
            If pRsGarantAvales!nmoneda = 1 Then
                If pRsGarantAvales!nTpoGarantia <> 4 Then
                    lnTotSoles = lnTotSoles + pRsGarantAvales!nRealizacion
                End If
            Else
                If pRsGarantAvales!nTpoGarantia <> 4 Then
                    lnTotDolar = lnTotDolar + pRsGarantAvales!nRealizacion
                End If
            End If
            
            Do While pRsGarantAvales!cNumGarant = lcNumGarantA And lnTotalDJ > 0 'And pRsGarantAvales!cPersCod = lcPersCod
                
                Word.Selection.TypeText Space(6) & Right(Space(5) & Format(pRsGarantAvales!nItem, "#,#0"), 5) & Space(2) & Replace(Left(UCase(Trim(pRsGarantAvales!cGarDjDescripcion)) & Space(30), 30), " ", " ") & Space(2) & Right(Space(5) & Format(pRsGarantAvales!nGarDJCantidad, "#,#0"), 5) & Space(12) & Right(Space(12) & Format(pRsGarantAvales!TotalDJ, "#,#0.00"), 12) & Space(2) & Replace(Left(UCase(pRsGarantAvales!Moneda) & Space(5), 5), " ", " ")
                Word.Selection.TypeParagraph

                If pRsGarantAvales!nmoneda = 1 Then
                    lnTotSoles = lnTotSoles + pRsGarantAvales!TotalDJ
                Else
                    lnTotDolar = lnTotDolar + pRsGarantAvales!TotalDJ
                End If

                pRsGarantAvales.MoveNext
                If pRsGarantAvales.EOF Then
                    Exit Do
                End If
                
            Loop

'*** PEAC 20080730

            If pRsGarantAvales.EOF Then
                Exit Do
            End If

            Word.Selection.TypeParagraph

            If pRsGarantAvales!cperscod = lcPersCod And pRsGarantAvales!cNumGarant = lcNumGarantA Then
                pRsGarantAvales.MoveNext
            End If
            
            'pRsGarantAvales.MoveNext
            If pRsGarantAvales.EOF Then
                Exit Do
            End If

        Loop
        
        Word.Selection.TypeParagraph
        Word.Selection.TypeText "======================================================================================"
        '----- pie de DOC
        Word.Selection.TypeParagraph
        Word.Selection.TypeText "TOTAL PATRIMONIO DECLARADO - SOLES :" & Right(Space(14) & Format(lnTotSoles, "#,#0.00"), 14)
        Word.Selection.TypeParagraph
        Word.Selection.TypeText "TOTAL PATRIMONIO DECLARADO - DOLAR :" & Right(Space(14) & Format(lnTotDolar, "#,#0.00"), 14)
        Word.Selection.TypeParagraph
        Word.Selection.TypeParagraph
        
        Word.Selection.TypeText "DEJO (AMOS) PLENA CONSTANCIA DE LOS BIENES DECLARADOS LO (LOS) OFREZCO (CEMOS) IRREVOCABLEMENTE  EN RESPALDO A  LOS" & _
                    " CREDITOS DIRECTOS E INDIRECTOS QUE TENGA A BIEN CONCEDERME (NOS) LA CAJA MUNICIPAL DE AHORRO Y CREDITO DE  MAYNAS." & _
                    " DECLARO (MOS) QUE LA INFORMACION PROPORCIONADA EN ESTE DOCUMENTO ES AUTENTICA Y AUTORIZO (AMOS) A LA CAJA MUNICIPAL" & _
                    " A REALIZAR SU VERIFICACION. ESTE DOCUMENTO TIENE CARACTER DE DECLARACION JURADA DE ACUERDO CON EL ARTICULO 179 DE" & _
                    " LA LEY No 26702 DECLARO ASI MISMO QUE EN CASO DE PROPORCIONAR INFORMACION FALSA ESTARE COMPRENDIDO DENTRO  DE LOS" & _
                    " ALCANCES  DEL  ART. 247 DEL  CODIGO PENAL QUE  ESTABLECE LA TIPIFICACION  DEL DELITO  FINANCIERO POR PROPORCIONAR" & _
                    " INFORMACION FALSA A UNA EMPRESA DEL SISTEMA FINANCIERO."
        Word.Selection.TypeParagraph
        Word.Selection.TypeParagraph
        
        Word.Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
        Word.Selection.TypeText Trim(pRsAgencia!Dist) & ", " & Format(pdFecSis, "dddd") & ", " & Day(pdFecSis) & " de " & Format(pdFecSis, "mmmm") & " de " & Year(pdFecSis)
        Word.Selection.TypeParagraph
        Word.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
        Word.Selection.TypeParagraph
        Word.Selection.TypeParagraph
        Word.Selection.TypeParagraph
        Word.Selection.TypeParagraph

            If Len(lcDniconyuge) > 0 Then
                Word.Selection.TypeText ".............................." & Space(22) & ".............................."
                Word.Selection.TypeParagraph
                Word.Selection.TypeText "DECLARANTE:" & Trim(lcPersNombre) & Space(2) & "CONYUGE: " & Trim(lcConyuge)
                Word.Selection.TypeParagraph
                Word.Selection.TypeText "DNI:" & Trim(lcDNI) & Space(44) & "DNI: " & Trim(lcDniconyuge)
                Word.Selection.TypeParagraph
            Else
                Word.Selection.TypeText ".............................."
                Word.Selection.TypeParagraph
                Word.Selection.TypeText "DECLARANTE:" & Trim(lcPersNombre)
                Word.Selection.TypeParagraph
                Word.Selection.TypeText "DNI:" & Trim(lcDNI)
                Word.Selection.TypeParagraph
            End If
        
        '---------- fin pie de DOC

        If pRsGarantAvales.EOF Then
            Exit Do
        End If

    Loop
    pRsGarantAvales.Close
    Set pRsGarantAvales = Nothing

End If

    'SELECCIONA TEXTO
    'Word.Selection.WholeStory
    'Word.Selection.Font.Size = 14
    
    ' VISIBLE
    Word.Visible = True
    
    Set Word = Nothing

End Function
''PEAC 20170621 - Este reporte se envio a gCredReportes
''WIOR 20130423 *************************************************************************************************
'Public Sub ImprimePagareCredPDF(ByVal psCtaCod As String, ByVal pnFormato As Integer, Optional ByVal psFecDes As String)
''***********Parametro psFecDes Agregado por PASI20131128 TI-ERS136-2013
'Dim oDoc  As cPDF
'Dim sLugar As String
'Dim sFecEmision As String
'Dim sFecVenc As String
'Dim sMoneda As String
'Dim sImporte As String
'Dim sImporteLetras As String
'Dim RelaGar As COMDPersona.DCOMPersonas ' PTI1 20170315
'Set RelaGar = New COMDPersona.DCOMPersonas  ' PTI1 20170315
'
'Dim nTasaMora As Double
'Dim nTasaInteres As Double
'
'Dim ssql As String
'Dim sCadImp As String
'Dim sNomAgencia As String
'Dim sEmision As String
'Dim lsCiudad As String
'
'Dim nGaran As Integer
'Dim nTitu As Integer
'Dim nCode As Integer
'Dim liPosicion As Integer
'
'Dim nTasaCompAnual As Double
'
'Dim R As ADODB.Recordset
'Dim RRelaCred As ADODB.Recordset
'Dim rsUbi As ADODB.Recordset
'Dim RsGarantes As New ADODB.Recordset
'
'Dim oDCred As COMDCredito.DCOMCredito
'Dim oFun As COMFunciones.FCOMCadenas
'Dim ObjCons As New COMDConstantes.DCOMAgencias
'Dim ObjGarantes As New COMDCredito.DCOMCredActBD
'
'sNomAgencia = ObjCons.NombreAgencia(Mid(psCtaCod, 4, 2))
'Set ObjCons = Nothing
'Set oDCred = New COMDCredito.DCOMCredito
'Set RRelaCred = oDCred.RecuperaRelacPers(psCtaCod)
'
'Set RsGarantes = oDCred.RecuperaGarantes(psCtaCod)
' Dim nCantGarant As Integer
'Dim sPersCodR As String 'PRT120170222, Agregó
'Dim RrelGar As ADODB.Recordset 'PRT120170222, Agregó
'Set oDCred = New COMDCredito.DCOMCredito
'Set R = oDCred.RecuperaDatosComunes(psCtaCod)
'
'Set rsUbi = oDCred.RecuperaUbigeo(Mid(psCtaCod, 4, 2))
'sEmision = rsUbi!cUbiGeoDescripcion
'
'
'lsCiudad = Trim(sEmision)
'liPosicion = InStr(lsCiudad, "(")
'
'If liPosicion > 0 Then
'    lsCiudad = Left(lsCiudad, liPosicion - 1)
'End If
'
'Set oDCred = Nothing
'Set oFun = New COMFunciones.FCOMCadenas
'
'Set oDoc = New cPDF
'
''Creación del Archivo
'oDoc.Author = gsCodUser
'oDoc.Creator = "SICMACT - Negocio"
'oDoc.Producer = "Caja Municipal de Ahorros y Crédito de Maynas S.A."
'oDoc.Subject = "Pagaré de Crédito Nº " & psCtaCod
'oDoc.Title = "Pagaré de Crédito Nº " & psCtaCod
'
'If Not oDoc.PDFCreate(App.path & "\Spooler\Pagare_" & psCtaCod & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
'    Exit Sub
'End If
'
'sMoneda = IIf(Mid(psCtaCod, 9, 1) = 1, gcPEN_PLURAL, "DOLARES") 'MARG ERS044-2016
'If Not (R.EOF And R.BOF) Then
''sFecVenc = Format(R!dVencPagare, "DD/MM/YYYY") -----Comentado por PTI1 090217-----
'    '*************************************************
'    'Modificado por PASI20131128 TI-ERS136-2013
'    'sFecEmision = Format(R!dVigencia, "DD/MM/YYYY") -----Comentado por PTI1 090217-----
'    If psFecDes = Empty Then
'        sFecEmision = Format(R!dVigencia, "DD/MM/YYYY")
'    Else
'        sFecEmision = Format(CDate(psFecDes), "DD/MM/YYYY")
'    End If
'    'END PASI*****************************************
'    'sImporte = Format(R!nMontoPagare, "#0.00") ---Comentado por PTI1 090217---
'    'nTasaMora = R!nTasaMora
'    nTasaMora = R!nTEAMora 'JUEZ 20140630
'    nTasaInteres = R!nTasaInteres
'    nTasaCompAnual = Format(((1 + nTasaInteres / 100) ^ (360 / 30) - 1) * 100, "#.00")
'    'sImporteLetras = NumLet(IIf(IsNull(R!nMontoPagare), 0, R!nMontoPagare)) & IIf(Mid(psCtaCod, 9, 1) = "2", "", " y " & IIf(InStr(1, R!nMontoPagare, ".") = 0, "00", Mid(IIf(IsNull(R!nMontoPagare), 0, R!nMontoPagare), InStr(1, IIf(IsNull(R!nMontoPagare), 0, R!nMontoPagare), ".") + 1, 2)) & "/100")
'    sImporteLetras = NumLet(IIf(IsNull(R!nMontoPagare), 0, R!nMontoPagare)) & IIf(Mid(psCtaCod, 9, 1) = "2", "", " y " & IIf(InStr(1, R!nMontoPagare, ".") = 0, "00", Left(Mid(IIf(IsNull(R!nMontoPagare), 0, R!nMontoPagare), InStr(1, IIf(IsNull(R!nMontoPagare), 0, R!nMontoPagare), ".") + 1, 2) & "00", 2)) & "/100") 'EJVG20130924
'Else
'    sFecVenc = ""
'    sFecEmision = ""
'    sImporte = ""
'    nTasaMora = 0
'    nTasaInteres = 0
'    nTasaCompAnual = 0
'    sImporteLetras = ""
'End If
'
'
'
'Rem oDoc.Fonts.Add "F1", "Arial", TrueType, Normal, WinAnsiEncoding
'Rem oDoc.Fonts.Add "F2", "Arial", TrueType, Bold, WinAnsiEncoding
'
'Rem oDoc.LoadImageFromFile App.Path & "\logo_cmacmaynas.bmp", "Logo"
'
''Tamaño de hoja A4
''oDoc.NewPage A4_Vertical
''------------------------------------------ COMENTADO PTI1 20170315------------------------------------
'Rem oDoc.WImage 75, 40, 35, 105, "Logo"
'Rem oDoc.WTextBox 63, 40, 15, 500, "PAGARE", "F2", 12, hCenter
'Rem 'oDoc.WTextBox 80, 40, 705, 520, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
'Rem oDoc.WTextBox 80, 40, 732, 520, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack 'JUEZ 20140630
'
'Rem oDoc.WTextBox 90, 60, 15, 160, "LUGAR DE EMISION", "F1", 10, hCenter, vMiddle, vbBlack, 1, vbBlack
'Rem oDoc.WTextBox 90, 220, 15, 160, "FECHA DE EMISION", "F1", 10, hCenter, vMiddle, vbBlack, 1, vbBlack
'Rem oDoc.WTextBox 90, 380, 15, 160, "NUMERO", "F1", 10, hCenter, vMiddle, vbBlack, 1, vbBlack
'
'Rem oDoc.WTextBox 105, 60, 15, 160, lsCiudad, "F2", 10, hCenter, vMiddle, vbBlack, 1, vbBlack
'Rem oDoc.WTextBox 105, 220, 15, 160, sFecEmision, "F2", 10, hCenter, vMiddle, vbBlack, 1, vbBlack
'Rem oDoc.WTextBox 105, 380, 15, 160, psCtaCod, "F2", 10, hCenter, vMiddle, vbBlack, 1, vbBlack
'
'Rem oDoc.WTextBox 120, 60, 15, 160, "FECHA DE VENCIMIENTO", "F1", 10, hCenter, vMiddle, vbBlack, 1, vbBlack
'Rem oDoc.WTextBox 120, 220, 15, 160, "MONEDA PAGARE", "F1", 10, hCenter, vMiddle, vbBlack, 1, vbBlack
'Rem oDoc.WTextBox 120, 380, 15, 160, "IMPORTE PAGARE", "F1", 10, hCenter, vMiddle, vbBlack, 1, vbBlack
'
'Rem oDoc.WTextBox 135, 60, 15, 160, sFecVenc, "F2", 10, hCenter, vMiddle, vbBlack, 1, vbBlack
'Rem oDoc.WTextBox 135, 220, 15, 160, sMoneda, "F2", 10, hCenter, vMiddle, vbBlack, 1, vbBlack
'Rem oDoc.WTextBox 135, 380, 15, 160, sImporte, "F2", 10, hCenter, vMiddle, vbBlack, 1, vbBlack
'
'Rem oDoc.WTextBox 150, 60, 20, 480, "Por éste PAGARE prometo (emos)  pagar incondicionalmente a la Orden de la CAJA MUNICIPAL DE AHORRO Y CREDITO DE MAYNAS S.A. (LA CAJA) la cantidad de:", "F1", 9, hjustify
'
'Rem oDoc.WTextBox 180, 60, 15, 320, UCase(sImporteLetras), "F2", 10, hCenter, vMiddle, vbBlack, 1, vbBlack
'Rem oDoc.WTextBox 180, 380, 15, 160, sMoneda, "F2", 10, hCenter, vMiddle, vbBlack, 1, vbBlack
'
'Rem oDoc.WTextBox 200, 60, 15, 450, "Importe a debitar en la siguiente cuenta de la Empresa que se indica:", "F1", 9, hjustify
'Rem oDoc.WTextBox 210, 60, 15, 450, "EMPRESA" & String(22, vbTab) & ": CAJA MUNICIPAL DE AHORRO Y CREDITO DE MAYNAS S.A.", "F2", 9
'Rem oDoc.WTextBox 220, 60, 15, 450, "OFICINA" & String(25, vbTab) & ": " & sNomAgencia, "F2", 9, hLeft
'Rem oDoc.WTextBox 230, 60, 15, 450, "NUMERO DE CUENTA" & String(2, vbTab) & ": " & psCtaCod, "F2", 9, hLeft
'Rem oDoc.WTextBox 240, 60, 15, 450, "D.C." & String(33, vbTab) & ": ", "F2", 9, hLeft
'
'Rem oDoc.WTextBox 260, 60, 20, 450, "Cláusulas Especiales:", "F1", 9, hjustify
''FRHU 20150305 CORREO
''oDoc.WTextBox 270, 75, 20, 470, "Este Pagaré debe ser pagado sólo en la misma moneda que expresa este título valor.", "F1", 9, hjustify
''oDoc.WTextBox 283, 75, 20, 470, "A su vencimiento, podrá ser prorrogado por su Tenedor, por el plazo que se señale en el reverso de este mismo documento, sin que sea necesario intervención alguna del obligado principal ni de los solidarios ó avales.", "F1", 9, hjustify
''oDoc.WTextBox 305, 75, 20, 470, "Desde su último vencimiento, su importe total y/o cuotas, generará los intereses compensatorios más moratorios a las tasas máximas autorizadas o permitidas a su último Tenedor.", "F1", 9, hjustify
''oDoc.WTextBox 327, 75, 20, 470, "De conformidad con lo dispuesto en los arts. 52 y 81 de la ley No 27287 el presente Pagaré ''NO REQUIERE PROTESTO'', pudiendo ejercitarse las acciones cambiarias por el solo mérito del vencimiento del plazo pactado, o de sus renovaciones.", "F1", 9, hjustify
''oDoc.WTextBox 358, 75, 20, 470, "El importe de éste Pagaré, podrá ser amortizado, renovándose por el monto del saldo.", "F1", 9, hjustify
'''oDoc.WTextBox 372, 75, 20, 470, "El importe de éste Pagaré y/o de sus cuotas, generará desde la emisión de éste Pagaré hasta la fecha de su respectivo vencimiento, un interés compensatorio a la tasa de " & nTasaCompAnual & " %.", "F1", 9, hjustify 'WIOR 20131025 'JUEZ 20140630 Comento
'''oDoc.WTextBox 372, 75, 20, 470, "El importe de éste Pagaré y/o de sus cuotas, generará desde la emisión de éste Pagaré hasta la fecha de su respectivo vencimiento, un interés compensatorio a la tasa de " & nTasaCompAnual & " % por año  y a partir de su vencimiento se cobrará adicionalmente un interés moratorio de " & nTasaMora & " %  diario.", "F1", 9, hjustify'WIOR 20131025 COMENTO
''oDoc.WTextBox 372, 75, 20, 470, "El importe de éste Pagaré y/o de sus cuotas, generará desde la emisión de éste Pagaré hasta la fecha de su respectivo vencimiento, un interés compensatorio a la tasa de " & nTasaCompAnual & " % por año y a partir de su vencimiento se cobrará adicionalmente un interés moratorio de " & nTasaMora & " % por año.", "F1", 9, hjustify 'JUEZ 20140630
''oDoc.WTextBox 403, 75, 20, 470, "Los pagos que correspondan, podrán ser verificados con cargo a la cuenta señalada en nuestra entidad ó de un Banco.", "F1", 9, hjustify
''oDoc.WTextBox 425, 75, 20, 470, "El  (los) emitente(s) y aval(es), autorizamos a retirar de mi (nuestra) cuenta(s) de Ahorro ó Plazo que en cualquier moneda mantenga(mos) en  LA CAJA  la suma para amortizar o pagar la presente obligación.", "F1", 9, hjustify
''oDoc.WTextBox 447, 75, 20, 470, "Los intereses compensatorios y moratorios podrán variar de acuerdo a la política institucional de LA CAJA o cuando el BCRP lo determine.", "F1", 9, hjustify
''oDoc.WTextBox 469, 75, 20, 470, "Este Pagaré  tiene naturaleza mercantil y se sujeta a las disposiciones de la Ley 27287 de Títulos Valores (Artículos 158º en adelante), la Ley de Bancos y al proceso ejecutivo señalado en el Código Procesal Civil en su caso.", "F1", 9, hjustify 'JUEZ 20140630
''
''oDoc.WTextBox 270, 60, 20, 20, "1.", "F1", 9, hjustify
''oDoc.WTextBox 283, 60, 20, 20, "2.", "F1", 9, hjustify
''oDoc.WTextBox 305, 60, 20, 20, "3.", "F1", 9, hjustify
''oDoc.WTextBox 327, 60, 20, 20, "4.", "F1", 9, hjustify
''oDoc.WTextBox 358, 60, 20, 20, "5.", "F1", 9, hjustify
''oDoc.WTextBox 372, 60, 20, 20, "6.", "F1", 9, hjustify
''oDoc.WTextBox 403, 60, 20, 20, "7.", "F1", 9, hjustify
''oDoc.WTextBox 425, 60, 20, 20, "8.", "F1", 9, hjustify
''oDoc.WTextBox 447, 60, 20, 20, "9.", "F1", 9, hjustify
''oDoc.WTextBox 469, 60, 20, 20, "10.", "F1", 9, hjustify 'JUEZ 20140630
'
'Rem oDoc.WTextBox 270, 75, 20, 470, "Este Pagaré debe ser pagado sólo en la misma moneda que expresa este título valor.", "F1", 9, hjustify
'Rem oDoc.WTextBox 283, 75, 20, 470, "Desde su último vencimiento, su importe total y/o cuotas, generará los intereses compensatorios más moratorios a las tasas máximas autorizadas o permitidas a su último Tenedor.", "F1", 9, hjustify
'Rem oDoc.WTextBox 305, 75, 20, 470, "De conformidad con lo dispuesto en los arts. 52 y 81 de la ley No 27287 el presente Pagaré ''NO REQUIERE PROTESTO'', pudiendo ejercitarse las acciones cambiarias por el solo mérito del vencimiento del plazo pactado, o de sus renovaciones.", "F1", 9, hjustify
'Rem oDoc.WTextBox 336, 75, 20, 470, "El importe de éste Pagaré, podrá ser amortizado parcial o totalmente, renovándose el mismo por el monto del saldo.", "F1", 9, hjustify
'Rem oDoc.WTextBox 358, 75, 20, 470, "El importe de éste Pagaré y/o de sus cuotas, generará desde la emisión de éste Pagaré hasta la fecha de su respectivo vencimiento, un interés compensatorio a la tasa de " & nTasaCompAnual & " % por año y a partir de su vencimiento se cobrará adicionalmente un interés moratorio de " & nTasaMora & " % por año.", "F1", 9, hjustify
'Rem oDoc.WTextBox 389, 75, 20, 470, "Los pagos que correspondan, podrán ser realizados a través de los canales de pago que LA CAJA pone a su disposición.", "F1", 9, hjustify
'Rem oDoc.WTextBox 411, 75, 20, 470, "El  (los) emitente(s) y aval(es), autorizamos a retirar de mi (nuestra) cuenta(s) de Ahorro ó Plazo que en cualquier moneda mantenga(mos) en  LA CAJA  la suma para amortizar o pagar la presente obligación.", "F1", 9, hjustify
'Rem oDoc.WTextBox 433, 75, 20, 470, "Los intereses compensatorios y moratorios podrán variar de acuerdo a lo convenido en el contrato de Mutuo suscrito por las partes.", "F1", 9, hjustify
'Rem oDoc.WTextBox 455, 75, 20, 470, "Este Pagaré  tiene naturaleza mercantil y se sujeta a las disposiciones de la Ley 27287 de Títulos Valores (Artículos 158º en adelante), la Ley de Bancos y al proceso ejecutivo señalado en el Código Procesal Civil en su caso.", "F1", 9, hjustify
'
'Rem oDoc.WTextBox 270, 60, 20, 20, "1.", "F1", 9, hjustify
'Rem oDoc.WTextBox 283, 60, 20, 20, "2.", "F1", 9, hjustify
'Rem oDoc.WTextBox 305, 60, 20, 20, "3.", "F1", 9, hjustify
'Rem oDoc.WTextBox 336, 60, 20, 20, "4.", "F1", 9, hjustify
'Rem oDoc.WTextBox 358, 60, 20, 20, "5.", "F1", 9, hjustify
'Rem oDoc.WTextBox 389, 60, 20, 20, "6.", "F1", 9, hjustify
'Rem oDoc.WTextBox 411, 60, 20, 20, "7.", "F1", 9, hjustify
'Rem oDoc.WTextBox 433, 60, 20, 20, "8.", "F1", 9, hjustify
'Rem oDoc.WTextBox 455, 60, 20, 20, "9.", "F1", 9, hjustify
'Rem 'FIN FRHU 20150305
'Rem Dim h As Integer 'JUEZ 20140630
'Rem h = 27
'
'Rem oDoc.WTextBox 473 + h, 60, 20, 450, "Emitente(s)", "F1", 9, hjustify
'
'Rem oDoc.WTextBox 485 + h, 300, 35, 50, "NOMBRE", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'Rem oDoc.WTextBox 520 + h, 300, 15, 50, "D.O.I.", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'Rem 'oDoc.WTextBox 535 + h, 300, 20, 50, "DOMICILIO", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'Rem oDoc.WTextBox 535 + h, 300, 35, 50, "DOMICILIO", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3 'FRHU 20150303 MEMO 491-2015
'
'Rem oDoc.WTextBox 485 + h, 350, 35, 205, "", "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'Rem oDoc.WTextBox 520 + h, 350, 15, 205, "", "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'Rem 'oDoc.WTextBox 535 + h, 350, 20, 205, "", "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'Rem oDoc.WTextBox 535 + h, 350, 35, 205, "", "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3 'FRHU 20150303 MEMO 491-2015
'
'Rem oDoc.WTextBox 580 + h, 300, 10, 255, String(40, "."), "F1", 9, hCenter
'Rem oDoc.WTextBox 590 + h, 300, 20, 255, "Firma Emitente", "F1", 9, hCenter
'
'Rem If Not (RRelaCred.EOF And RRelaCred.BOF) Then
'    Rem Do Until RRelaCred.EOF
'        Rem If RRelaCred!nConsValor = gColRelPersTitular Then
'
'            Rem oDoc.WTextBox 485 + h, 45, 35, 50, "NOMBRE", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'            Rem oDoc.WTextBox 520 + h, 45, 15, 50, "D.O.I.", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'            Rem 'oDoc.WTextBox 535 + h, 45, 20, 50, "DOMICILIO", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'            Rem oDoc.WTextBox 535 + h, 45, 35, 50, "DOMICILIO", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3 'FRHU 20150303 MEMO 491-2015
'
'            Rem oDoc.WTextBox 485 + h, 95, 35, 205, QuitarCaracter(PstaNombre(RRelaCred!cPersNombre), "-*/\._"), "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'            Rem oDoc.WTextBox 520 + h, 95, 15, 205, IIf(IsNull(RRelaCred!DNI), RRelaCred!Ruc, RRelaCred!DNI), "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'            Rem 'oDoc.WTextBox 535 + h, 95, 20, 205, RRelaCred!cPersDireccDomicilio, "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 4
'            Rem oDoc.WTextBox 535 + h, 95, 35, 205, RRelaCred!cPersDireccDomicilio, "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 4 'FRHU 20150303 MEMO 491-2015
'
'            Rem oDoc.WTextBox 580 + h, 45, 10, 255, String(40, "."), "F1", 9, hCenter
'            Rem oDoc.WTextBox 590 + h, 45, 15, 255, "Firma Emitente", "F1", 9, hCenter
'            Rem nTitu = 1
'
'        Rem ElseIf RRelaCred!nConsValor = gColRelPersConyugue Or RRelaCred!nConsValor = gColRelPersCodeudor Then
'            Rem oDoc.WTextBox 485 + h, 350, 35, 205, QuitarCaracter(PstaNombre(RRelaCred!cPersNombre), "-*/\._"), "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'            Rem oDoc.WTextBox 520 + h, 350, 15, 205, IIf(IsNull(RRelaCred!DNI), RRelaCred!Ruc, RRelaCred!DNI), "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'            Rem 'oDoc.WTextBox 535 + h, 350, 20, 205, RRelaCred!cPersDireccDomicilio, "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 4
'            Rem oDoc.WTextBox 535 + h, 350, 35, 205, RRelaCred!cPersDireccDomicilio, "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 4 'FRHU 20150303 MEMO 491-2015
'            Rem nCode = 1
'            Rem Exit Do
'        Rem End If
'        Rem RRelaCred.MoveNext
'    Rem Loop
'Rem End If
'
'Rem oDoc.WTextBox 610 + h, 60, 20, 450, "Aval(es)", "F1", 9, hjustify
'
'Rem oDoc.WTextBox 622 + h, 45, 35, 50, "NOMBRE", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'Rem oDoc.WTextBox 657 + h, 45, 15, 50, "D.O.I.", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'Rem 'oDoc.WTextBox 672 + h, 45, 20, 50, "DOMICILIO", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'Rem oDoc.WTextBox 672 + h, 45, 35, 50, "DOMICILIO", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3 'FRHU 20150303 MEMO 491-2015
'
'Rem oDoc.WTextBox 622 + h, 95, 35, 205, "", "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'Rem oDoc.WTextBox 657 + h, 95, 15, 205, "", "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'Rem 'oDoc.WTextBox 672 + h, 95, 20, 205, "", "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 4
'Rem oDoc.WTextBox 672 + h, 95, 35, 205, "", "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 4 'FRHU 20150303 MEMO 491-2015
'
'Rem oDoc.WTextBox 717 + h, 45, 10, 255, String(40, "."), "F1", 9, hCenter
'Rem oDoc.WTextBox 727 + h, 45, 15, 255, "Firma Aval", "F1", 9, hCenter
'
'
'Rem oDoc.WTextBox 622 + h, 300, 35, 50, "NOMBRE", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'Rem oDoc.WTextBox 657 + h, 300, 15, 50, "D.O.I.", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'Rem 'oDoc.WTextBox 672 + h, 300, 20, 50, "DOMICILIO", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'Rem oDoc.WTextBox 672 + h, 300, 35, 50, "DOMICILIO", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3 'FRHU 20150303 MEMO 491-2015
'
'Rem oDoc.WTextBox 622 + h, 350, 35, 205, "", "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'Rem oDoc.WTextBox 657 + h, 350, 15, 205, "", "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'Rem 'oDoc.WTextBox 672 + h, 350, 20, 205, "", "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'Rem oDoc.WTextBox 672 + h, 350, 35, 205, "", "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3 'FRHU 20150303 MEMO 491-2015
'
'Rem oDoc.WTextBox 717 + h, 300, 10, 255, String(40, "."), "F1", 9, hCenter
'Rem oDoc.WTextBox 727 + h, 300, 20, 255, "Firma Aval", "F1", 9, hCenter
'
'Rem Dim nCantGarant As Integer
'
'Rem nCantGarant = RsGarantes.RecordCount
'Rem nGaran = 0
'
'Rem If Not (RsGarantes.EOF And RsGarantes.BOF) Then
'        Rem While Not RsGarantes.EOF
'            Rem If nGaran = 0 Then
'                Rem oDoc.WTextBox 622 + h, 95, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'                Rem oDoc.WTextBox 657 + h, 95, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'                Rem 'oDoc.WTextBox 672 + h, 95, 20, 205, RsGarantes!cPersDireccDomicilio, "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 4
'                Rem oDoc.WTextBox 672 + h, 95, 35, 205, RsGarantes!cPersDireccDomicilio, "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 4 'FRHU 20150303 MEMO 491-2015
'            Rem ElseIf nGaran = 1 Then
'                Rem oDoc.WTextBox 622 + h, 350, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'                Rem oDoc.WTextBox 657 + h, 350, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'                Rem 'oDoc.WTextBox 672 + h, 350, 20, 205, RsGarantes!cPersDireccDomicilio, "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 4
'                Rem oDoc.WTextBox 672 + h, 350, 35, 205, RsGarantes!cPersDireccDomicilio, "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 4 'FRHU 20150303 MEMO 491-2015
'            Rem End If
'            Rem nGaran = nGaran + 1
'            Rem RsGarantes.MoveNext
'        Rem Wend
'    Rem Else
'Rem End If
'
'
'Rem Dim sPiePagina As String
'
''FRHU 20150305 CORREO
''sPiePagina = "Oficina Principal: Jr. Próspero No  791 -  Iquitos - Tel.  (065) 221-256, 223-323; " & _
''            "Ag. Calle Arequipa: Ca Arequipa 428 - Tel.  (065) 222-765; Ag. Pucallpa: Jr. Ucayali  No  850 - 852 - Tel. (061) 593-671; " & _
''            "Ag. Huánuco: Jr. 28 de Julio  No  944 - Tel. (062) 513-017, 514-340; Ag. Belén: Av. Grau 1260   Tel. (065) 268453; " & _
''            "Ag. Yurimaguas: Ca. Simón Bolívar 113 Tel (065)  351235; Ag. Tingo María: Av. Antonio Raymondi 246 - Tel. (062) 561149; " & _
''            "Ag. Tarapoto: Jr San Martín 205 - Tel. (042) 526078;  Ag. Requena: Ca. Malecón Bolognesi 305 - Tel. (065) 412010; " & _
''            "Ag. Cajamarca: Jr. Amalia Puga 419 - Tel. (076) 344896; Of.  Esp. Aguaytía: Simón Bolívar Nº 298 - Tel. (061) 481566; " & _
''            "Ag. Cerro de Pasco; Plaza Carrión Nº 191- Tel: (063) 421414; Agencia Punchana Av. 28 de Julio 829 - Tel (065) 252277."
'Rem sPiePagina = "Oficina Principal: Jr. Próspero No  791 -  Iquitos ;" & _
'             REM "Ag. Calle Arequipa: Ca Arequipa Nº 428;  Agencia Punchana Av. 28 de Julio 829 - Iquitos;" & _
'             REM "Ag. Belén: Av. Grau Nº 1260 - Iquitos; Ag. San Juan Bautista- Avda. Abelardo Quiñones Nº 2670- Iquitos;" & _
'             REM "Ag. Pucallpa: Jr. Ucayali  No  850 - 852 ; Ag. Huánuco: Jr. General Prado No   836;" & _
'             REM "Ag. Yurimaguas: Ca. Simón Bolívar Nº 113; Ag. Tingo María: Av. Antonio Raymondi  Nº 246 ;" & _
'             REM "Ag. Tarapoto: Jr San Martín Nº 205 ;  Ag. Requena: Calle San Francisco Mz 28 Lt 07;" & _
'             REM "Ag. Cajamarca: Jr. Amalia Puga  Nº 417; Ag.  Aguaytía- Jr. Rio Negro Nº 259;" & _
'             REM "Ag. Cerro de Pasco; Plaza Carrión Nº 191; Ag. Minka: Ciudad Comercial Minka - Av. Argentina Nº 3093- Local 230 - Callao."
'Rem 'FIN FRHU 20150305
'Rem oDoc.WTextBox 740 + h, 45, 15, 510, sPiePagina, "F1", 7, hjustify
'
'Rem If nCantGarant > 2 Then
'    Rem oDoc.NewPage A4_Vertical
'
'    Rem oDoc.WImage 75, 40, 35, 105, "Logo"
'    Rem oDoc.WTextBox 63, 40, 15, 500, "PAGARE", "F2", 12, hCenter
'    Rem oDoc.WTextBox 80, 40, 705, 520, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
'
'    Rem oDoc.WTextBox 740, 45, 15, 510, sPiePagina, "F1", 7, hjustify
'
'    Rem oDoc.WTextBox 88, 60, 20, 450, "Aval(es)", "F1", 9, hjustify
'    Rem oDoc.WTextBox 100, 45, 35, 50, "NOMBRE", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'    Rem oDoc.WTextBox 135, 45, 15, 50, "D.O.I.", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'    Rem 'oDoc.WTextBox 150, 45, 20, 50, "DOMICILIO", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'    Rem oDoc.WTextBox 150, 45, 35, 50, "DOMICILIO", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3 'FRHU 20150303 MEMO 491-2015
'
'    Rem oDoc.WTextBox 100, 95, 35, 205, "", "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'    Rem oDoc.WTextBox 135, 95, 15, 205, "", "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'    Rem 'oDoc.WTextBox 150, 95, 20, 205, "", "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 4
'    Rem oDoc.WTextBox 150, 95, 35, 205, "", "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 4 'FRHU 20150303 MEMO 491-2015
'
'    Rem oDoc.WTextBox 195, 45, 10, 255, String(40, "."), "F1", 9, hCenter
'    Rem oDoc.WTextBox 205, 45, 15, 255, "Firma Aval", "F1", 9, hCenter
'
'
'    Rem oDoc.WTextBox 100, 300, 35, 50, "NOMBRE", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'    Rem oDoc.WTextBox 135, 300, 15, 50, "D.O.I.", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'    Rem 'oDoc.WTextBox 150, 300, 20, 50, "DOMICILIO", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'    Rem oDoc.WTextBox 150, 300, 35, 50, "DOMICILIO", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3 'FRHU 20150303 MEMO 491-2015
'
'    Rem oDoc.WTextBox 100, 350, 35, 205, "", "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'    Rem oDoc.WTextBox 135, 350, 15, 205, "", "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'    Rem 'oDoc.WTextBox 150, 350, 20, 205, "", "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'    Rem oDoc.WTextBox 150, 350, 35, 205, "", "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3 'FRHU 20150303 MEMO 491-2015
'
'    Rem oDoc.WTextBox 195, 300, 10, 255, String(40, "."), "F1", 9, hCenter
'    Rem oDoc.WTextBox 205, 300, 20, 255, "Firma Aval", "F1", 9, hCenter
'
'    Rem nGaran = 0
'    Rem RsGarantes.MoveFirst
'    Rem If Not (RsGarantes.EOF And RsGarantes.BOF) Then
'            Rem While Not RsGarantes.EOF
'                Rem If nGaran = 2 Then
'                    Rem oDoc.WTextBox 100, 95, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'                    Rem oDoc.WTextBox 135, 95, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'                    Rem 'oDoc.WTextBox 150, 95, 20, 205, RsGarantes!cPersDireccDomicilio, "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 4
'                    Rem oDoc.WTextBox 150, 95, 35, 205, RsGarantes!cPersDireccDomicilio, "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 4 'FRHU 20150303 MEMO 491-2015
'                Rem ElseIf nGaran = 3 Then
'                    Rem oDoc.WTextBox 100, 350, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'                    Rem oDoc.WTextBox 135, 350, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'                    Rem 'oDoc.WTextBox 150, 350, 20, 205, RsGarantes!cPersDireccDomicilio, "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 4
'                    Rem oDoc.WTextBox 150, 350, 35, 205, RsGarantes!cPersDireccDomicilio, "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 4 'FRHU 20150303 MEMO 491-2015
'                Rem End If
'                Rem nGaran = nGaran + 1
'                Rem RsGarantes.MoveNext
'            Rem Wend
'        Rem Else
'    Rem End If
'
'    Rem If nCantGarant > 4 Then
'        Rem oDoc.WTextBox 225, 60, 20, 450, "Aval(es)", "F1", 9, hjustify
'        Rem oDoc.WTextBox 237, 45, 35, 50, "NOMBRE", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'        Rem oDoc.WTextBox 272, 45, 15, 50, "D.O.I.", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'        Rem 'oDoc.WTextBox 287, 45, 20, 50, "DOMICILIO", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'        Rem oDoc.WTextBox 287, 45, 35, 50, "DOMICILIO", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3 'FRHU 20150303 MEMO 491-2015
'
'        Rem oDoc.WTextBox 237, 95, 35, 205, "", "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'        Rem oDoc.WTextBox 272, 95, 15, 205, "", "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'        Rem 'oDoc.WTextBox 287, 95, 20, 205, "", "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 4
'        Rem oDoc.WTextBox 287, 95, 35, 205, "", "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 4 'FRHU 20150303 MEMO 491-2015
'
'        Rem oDoc.WTextBox 332, 45, 10, 255, String(40, "."), "F1", 9, hCenter
'        Rem oDoc.WTextBox 342, 45, 15, 255, "Firma Aval", "F1", 9, hCenter
'
'
'        Rem oDoc.WTextBox 237, 300, 35, 50, "NOMBRE", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'        Rem oDoc.WTextBox 272, 300, 15, 50, "D.O.I.", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'        Rem 'oDoc.WTextBox 287, 300, 20, 50, "DOMICILIO", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'        Rem oDoc.WTextBox 287, 300, 35, 50, "DOMICILIO", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3 'FRHU 20150303 MEMO 491-2015
'
'        Rem oDoc.WTextBox 237, 350, 35, 205, "", "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'        Rem oDoc.WTextBox 272, 350, 15, 205, "", "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'        Rem 'oDoc.WTextBox 287, 350, 20, 205, "", "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'        Rem oDoc.WTextBox 287, 350, 35, 205, "", "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3 'FRHU 20150303 MEMO 491-2015
'
'        Rem oDoc.WTextBox 332, 300, 10, 255, String(40, "."), "F1", 9, hCenter
'        Rem oDoc.WTextBox 342, 300, 20, 255, "Firma Aval", "F1", 9, hCenter
'
'        Rem nGaran = 0
'        Rem RsGarantes.MoveFirst
'        Rem If Not (RsGarantes.EOF And RsGarantes.BOF) Then
'                Rem While Not RsGarantes.EOF
'                    Rem If nGaran = 4 Then
'                        Rem oDoc.WTextBox 237, 95, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'                        Rem oDoc.WTextBox 272, 95, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'                        Rem 'oDoc.WTextBox 287, 95, 20, 205, RsGarantes!cPersDireccDomicilio, "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 4
'                        Rem oDoc.WTextBox 287, 95, 35, 205, RsGarantes!cPersDireccDomicilio, "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 4 'FRHU 20150303 MEMO 491-2015
'                    Rem ElseIf nGaran = 5 Then
'                        Rem oDoc.WTextBox 237, 350, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'                        Rem oDoc.WTextBox 272, 350, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'                        Rem 'oDoc.WTextBox 287, 350, 20, 205, RsGarantes!cPersDireccDomicilio, "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 4
'                        Rem oDoc.WTextBox 287, 350, 35, 205, RsGarantes!cPersDireccDomicilio, "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 4 'FRHU 20150303 MEMO 491-2015
'                    Rem End If
'                    Rem nGaran = nGaran + 1
'                    Rem RsGarantes.MoveNext
'                Rem Wend
'            Rem Else
'        Rem End If
'    Rem End If
'
'    Rem If nCantGarant > 6 Then
'        Rem oDoc.WTextBox 362, 60, 20, 450, "Aval(es)", "F1", 9, hjustify
'        Rem oDoc.WTextBox 374, 45, 35, 50, "NOMBRE", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'        Rem oDoc.WTextBox 409, 45, 15, 50, "D.O.I.", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'        Rem oDoc.WTextBox 424, 45, 20, 50, "DOMICILIO", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'
'        Rem oDoc.WTextBox 374, 95, 35, 205, "", "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'        Rem oDoc.WTextBox 409, 95, 15, 205, "", "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'        Rem oDoc.WTextBox 424, 95, 20, 205, "", "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 4
'
'        Rem oDoc.WTextBox 469, 45, 10, 255, String(40, "."), "F1", 9, hCenter
'        Rem oDoc.WTextBox 479, 45, 15, 255, "Firma Aval", "F1", 9, hCenter
'
'
'        Rem oDoc.WTextBox 374, 300, 35, 50, "NOMBRE", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'        Rem oDoc.WTextBox 409, 300, 15, 50, "D.O.I.", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'        Rem oDoc.WTextBox 424, 300, 20, 50, "DOMICILIO", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'
'        Rem oDoc.WTextBox 374, 350, 35, 205, "", "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'        Rem oDoc.WTextBox 409, 350, 15, 205, "", "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'        Rem oDoc.WTextBox 424, 350, 20, 205, "", "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'
'        Rem oDoc.WTextBox 469, 300, 10, 255, String(40, "."), "F1", 9, hCenter
'        Rem oDoc.WTextBox 479, 300, 20, 255, "Firma Aval", "F1", 9, hCenter
'
'        Rem nGaran = 0
'        Rem RsGarantes.MoveFirst
'        Rem If Not (RsGarantes.EOF And RsGarantes.BOF) Then
'                Rem While Not RsGarantes.EOF
'                    Rem If nGaran = 6 Then
'                        Rem oDoc.WTextBox 374, 95, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'                        Rem oDoc.WTextBox 409, 95, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'                        Rem oDoc.WTextBox 424, 95, 20, 205, RsGarantes!cPersDireccDomicilio, "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 4
'                    Rem ElseIf nGaran = 7 Then
'                        Rem oDoc.WTextBox 374, 350, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'                        Rem oDoc.WTextBox 409, 350, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'                        Rem oDoc.WTextBox 424, 350, 20, 205, RsGarantes!cPersDireccDomicilio, "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 4
'
'                    Rem End If
'                    Rem nGaran = nGaran + 1
'                    Rem RsGarantes.MoveNext
'                Rem Wend
'            Rem Else
'        Rem End If
'    Rem End If
'
'    Rem If nCantGarant > 8 Then
'
'        Rem oDoc.WTextBox 499, 60, 20, 450, "Aval(es)", "F1", 9, hjustify
'        Rem oDoc.WTextBox 511, 45, 35, 50, "NOMBRE", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'        Rem oDoc.WTextBox 546, 45, 15, 50, "D.O.I.", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'        Rem oDoc.WTextBox 561, 45, 20, 50, "DOMICILIO", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'
'        Rem oDoc.WTextBox 511, 95, 35, 205, "", "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'        Rem oDoc.WTextBox 546, 95, 15, 205, "", "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'        Rem oDoc.WTextBox 561, 95, 20, 205, "", "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 4
'
'        Rem oDoc.WTextBox 606, 45, 10, 255, String(40, "."), "F1", 9, hCenter
'        Rem oDoc.WTextBox 616, 45, 15, 255, "Firma Aval", "F1", 9, hCenter
'
'
'        Rem oDoc.WTextBox 511, 300, 35, 50, "NOMBRE", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'        Rem oDoc.WTextBox 546, 300, 15, 50, "D.O.I.", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'        Rem oDoc.WTextBox 561, 300, 20, 50, "DOMICILIO", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'
'        Rem oDoc.WTextBox 511, 350, 35, 205, "", "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'        Rem oDoc.WTextBox 546, 350, 15, 205, "", "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'        Rem oDoc.WTextBox 561, 350, 20, 205, "", "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'
'        Rem oDoc.WTextBox 606, 300, 10, 255, String(40, "."), "F1", 9, hCenter
'        Rem oDoc.WTextBox 616, 300, 20, 255, "Firma Aval", "F1", 9, hCenter
'
'        Rem nGaran = 0
'        Rem RsGarantes.MoveFirst
'        Rem If Not (RsGarantes.EOF And RsGarantes.BOF) Then
'                Rem While Not RsGarantes.EOF
'                    Rem If nGaran = 8 Then
'                        Rem oDoc.WTextBox 511, 95, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'                        Rem oDoc.WTextBox 546, 95, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'                        Rem oDoc.WTextBox 561, 95, 20, 205, RsGarantes!cPersDireccDomicilio, "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 4
'                    Rem ElseIf nGaran = 9 Then
'                        Rem oDoc.WTextBox 511, 350, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'                        Rem oDoc.WTextBox 546, 350, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'                        Rem oDoc.WTextBox 561, 350, 20, 205, RsGarantes!cPersDireccDomicilio, "F2", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 4
'
'                    Rem End If
'                    Rem nGaran = nGaran + 1
'                    Rem RsGarantes.MoveNext
'                Rem Wend
'            Rem Else
'        Rem End If
'    Rem End If
'Rem End If
''------------------------------------------ END COMENTADO PTI1 20170315 ------------------------------------
''######################## PTI1 20170315 ACTUALIZACION FORMATO ########################################
'
'oDoc.Fonts.Add "F1", "arial narrow", TrueType, Normal, WinAnsiEncoding
'oDoc.Fonts.Add "F2", "arial narrow", TrueType, Bold, WinAnsiEncoding
'
'oDoc.LoadImageFromFile App.path & "\Logo_2015.jpg", "Logo"
'
'oDoc.NewPage A4_Vertical
'
'
'oDoc.WImage 50, 494, 35, 73, "Logo"
'oDoc.WTextBox 30, 40, 15, 500, "PAGARÉ", "F2", 12, hCenter
'oDoc.WTextBox 60, 45, 15, 175, "LUGAR DE EMISIÓN", "F1", 10, hCenter, vMiddle, vbBlack, 1, vbBlack
'oDoc.WTextBox 60, 220, 15, 160, "FECHA DE EMISIÓN", "F1", 10, hCenter, vMiddle, vbBlack, 1, vbBlack
'oDoc.WTextBox 60, 380, 15, 187, "NÚMERO", "F1", 10, hCenter, vMiddle, vbBlack, 1, vbBlack
'
'oDoc.WTextBox 75, 45, 15, 175, lsCiudad, "F2", 10, hCenter, vMiddle, vbBlack, 1, vbBlack
''oDoc.WTextBox 75, 220, 15, 160, sFecEmision, "F2", 10, hCenter, vMiddle, vbBlack, 1, vbBlack 'comentado PTI1 29-03-2017
'oDoc.WTextBox 75, 380, 15, 187, psCtaCod, "F2", 10, hCenter, vMiddle, vbBlack, 1, vbBlack
'
'oDoc.WTextBox 90, 45, 15, 175, "FECHA DE VENCIMIENTO", "F1", 10, hCenter, vMiddle, vbBlack, 1, vbBlack
'oDoc.WTextBox 90, 220, 15, 160, "MONEDA PAGARÉ", "F1", 10, hCenter, vMiddle, vbBlack, 1, vbBlack
'oDoc.WTextBox 90, 380, 15, 187, "IMPORTE PAGARÉ", "F1", 10, hCenter, vMiddle, vbBlack, 1, vbBlack
'
'oDoc.WTextBox 105, 45, 15, 175, sFecVenc, "F2", 10, hCenter, vMiddle, vbBlack, 1, vbBlack
'oDoc.WTextBox 105, 220, 15, 160, sMoneda, "F2", 10, hCenter, vMiddle, vbBlack, 1, vbBlack
'oDoc.WTextBox 105, 380, 15, 187, sImporte, "F2", 10, hCenter, vMiddle, vbBlack, 1, vbBlack
''----------------PTI1----------------------------------------------
'oDoc.WTextBox 130, 45, 20, 480, "Por este", "F1", 11, hjustify
'oDoc.WTextBox 131, 80, 20, 480, "PAGARÉ", "F2", 10, hjustify
'oDoc.WTextBox 130, 117, 20, 480, "prometo/prometemos  pagar solidariamente e incondicionalmente a la orden de la", "F1", 11, hjustify
'oDoc.WTextBox 131, 447, 20, 480, "CAJA MUNICIPAL DE AHORRO", "F2", 10, hjustify
'oDoc.WTextBox 141, 45, 30, 480, "Y CRÉDITO DE MAYNAS S.A.", "F2", 10, hjustify
'oDoc.WTextBox 140, 158, 30, 480, ", con R.U.C N° 20103845328, en adelante ", "F1", 11, hjustify
'oDoc.WTextBox 141, 330, 30, 480, "LA CAJA", "F2", 10, hjustify
'oDoc.WTextBox 140, 363, 30, 480, ", en cualquiera de sus oficinas a nivel nacional, o a", "F1", 11, hjustify
'oDoc.WTextBox 150, 45, 30, 480, "quien", "F1", 11, hjustify
'oDoc.WTextBox 151, 68, 30, 480, "LA CAJA", "F2", 10, hjustify
'oDoc.WTextBox 150, 104, 30, 480, "hubiera endosado el presente título valor", "F1", 11, hjustify
'oDoc.WTextBox 150, 267, 30, 600, ", la suma de__________________________________________________", "F1", 11, hjustify
'oDoc.WTextBox 160, 45, 30, 600, "_________________________________________________", "F1", 11, hjustify
'oDoc.WTextBox 160, 290, 30, 480, ", importe de dinero que expresamente declaro/declaramos adeudar a", "F1", 11, hjustify
'oDoc.WTextBox 171, 45, 30, 480, "LA CAJA", "F2", 10, hjustify
'oDoc.WTextBox 170, 80, 30, 500, "y que me(nos) obligo/obligamos a pagar en la misma moneda antes expresada en la fecha de vencimiento consignada. ", "F1", 11, hjustify
''-------------------------------------------------------------------------------------------
'oDoc.WTextBox 190, 45, 30, 540, "Queda " & String(0.5, vbTab) & "expresamente" & String(0.55, vbTab) & " estipulado" & String(0.55, vbTab) & " que" & String(0.55, vbTab) & " el" & String(0.55, vbTab) & " importe" & String(0.55, vbTab) & " de" & String(0.55, vbTab) & " este" & String(0.55, vbTab) & " Pagaré " & String(0.55, vbTab) & "devengará" & String(0.55, vbTab) & " desde" & String(0.55, vbTab) & " su" & String(0.55, vbTab) & " fecha" & String(0.55, vbTab) & " de" & String(0.55, vbTab) & " emisión " & String(0.55, vbTab) & "hasta" & String(0.55, vbTab) & " la" & String(0.55, vbTab) & " fecha" & String(0.55, vbTab) & " de " & String(0.55, vbTab) & "", "F1", 11, hjustify
'oDoc.WTextBox 200, 45, 30, 540, "su" & String(0.55, vbTab) & " vencimiento" & String(0.55, vbTab) & " un" & String(0.55, vbTab) & " interés" & String(0.55, vbTab) & " compensatorio" & String(0.55, vbTab) & " a" & String(0.55, vbTab) & " una" & String(0.55, vbTab) & " tasa" & String(0.55, vbTab) & " efectiva" & String(0.55, vbTab) & " anual" & String(0.55, vbTab) & " del" & String(0.55, vbTab) & "", "F1", 11, hjustify
'oDoc.WTextBox 200, 394, 30, 520, "y" & String(0.55, vbTab) & " a" & String(0.55, vbTab) & "  partir" & String(0.55, vbTab) & "  de" & String(0.55, vbTab) & " su" & String(0.55, vbTab) & "  vencimiento" & String(0.55, vbTab) & " se" & String(0.55, vbTab) & " cobrará", "F1", 11, hjustify
'oDoc.WTextBox 201, 358, 30, 520, "" & nTasaCompAnual & "%", "F2", 10, hjustify
'oDoc.WTextBox 210, 45, 30, 520, "adicionalmente" & String(0.54, vbTab) & "un" & String(0.54, vbTab) & "interés" & String(0.54, vbTab) & "moratorio" & String(0.54, vbTab) & "a" & String(0.54, vbTab) & " una" & String(0.54, vbTab) & " tasa" & String(0.54, vbTab) & " efectiva" & String(0.54, vbTab) & " anual" & String(0.54, vbTab) & " del" & String(0.54, vbTab) & "", "F1", 11, hjustify
'oDoc.WTextBox 211, 320, 30, 520, "" & nTasaMora & ".00%.", "F2", 10, hjustify
'oDoc.WTextBox 210, 358, 30, 400, " Ambas" & String(0.54, vbTab) & "tasas" & String(0.54, vbTab) & "de" & String(0.54, vbTab) & "interés" & String(0.54, vbTab) & " continuarán" & String(0.54, vbTab) & "devengándose", "F1", 11, hjustify
'oDoc.WTextBox 220, 45, 30, 520, "por todo el tiempo que demore el pago de la presente obligación.", "F1", 11, hjustify
''-------------------------------------------------------------------------------------------
'oDoc.WTextBox 240, 45, 30, 515, "Asimismo " & String(0.51, vbTab) & " autorizo(amos) " & String(0.51, vbTab) & " de " & String(0.51, vbTab) & " manera" & String(0.51, vbTab) & " expresa " & String(0.51, vbTab) & " el cobro" & String(0.51, vbTab) & " de penalidades, seguros, gastos " & String(0.51, vbTab) & " notariales, de " & String(0.51, vbTab) & " cobranza judicial y " & String(300, vbTab) & "", "F1", 11, hjustify
'oDoc.WTextBox 250, 45, 30, 520, String(2, vbTab) & " extrajudicial, y en" & String(0.54, vbTab) & " general" & String(0.54, vbTab) & "los gastos" & String(0.54, vbTab) & "y comisiones que pudiéramos adeudar derivados del crédito representado en este", "F1", 11, hjustify
'oDoc.WTextBox 250, 535, 30, 520, "Pagaré,", "F1", 11, hjustify
'oDoc.WTextBox 260, 45, 30, 540, "y que se pudieran generar desde la fecha de emisión del presente Pagaré hasta la cancelación total de la presente obligación,", "F1", 11, hjustify
'oDoc.WTextBox 260, 554, 30, 540, "sin", "F1", 11, hjustify
'oDoc.WTextBox 270, 45, 30, 540, "que" & String(0.55, vbTab) & "sea necesario" & String(0.55, vbTab) & " requerimiento" & String(0.55, vbTab) & " alguno" & String(0.55, vbTab) & " de" & String(0.55, vbTab) & " pago " & String(0.55, vbTab) & "para", "F1", 11, hjustify
'oDoc.WTextBox 270, 278, 30, 540, "constituirme/constituirnos" & String(1, vbTab) & " en" & String(2, vbTab) & " mora," & String(0.54, vbTab) & " pues" & String(2, vbTab) & " es" & String(0.54, vbTab) & " entendido" & String(0.54, vbTab) & " que" & String(0.54, vbTab) & " ésta se ", "F1", 11, hjustify
'oDoc.WTextBox 280, 45, 30, 540, "producirá de modo automático por el solo hecho del vencimiento de éste Pagaré.", "F1", 11, hjustify
''-------------------------------------------------------------------------------------------
'oDoc.WTextBox 300, 45, 30, 540, "Expresamente" & String(0.55, vbTab) & " acepto(amos) toda" & String(1, vbTab) & " variación" & String(1, vbTab) & " de" & String(1, vbTab) & " las " & String(0.5, vbTab) & "tasas" & String(0.5, vbTab) & " de interés, dentro de los límites legales autorizados, las mismas que se ", "F1", 11, hjustify
'oDoc.WTextBox 310, 45, 30, 540, "aplicarán" & String(0.55, vbTab) & " luego" & String(0.55, vbTab) & " de" & String(0.55, vbTab) & " la" & String(0.55, vbTab) & " comunicación" & String(0.55, vbTab) & " efectuada" & String(0.55, vbTab) & " por" & String(0.55, vbTab) & " la ", "F1", 11, hjustify
'oDoc.WTextBox 311, 274, 30, 480, "LA CAJA", "F2", 10, hjustify
'oDoc.WTextBox 310, 308, 30, 480, ", conforme a ley. Se" & String(0.55, vbTab) & " deja constancia que el presente Pagaré " & """" & "no", "F1", 11, hjustify
'oDoc.WTextBox 320, 45, 30, 540, "requiere" & String(0.6, vbTab) & " ser" & String(0.55, vbTab) & " protestado" & """" & " por" & String(1.4, vbTab) & " falta" & String(1.4, vbTab) & " de" & String(1.4, vbTab) & " pago, procediendo" & String(1.4, vbTab) & "su ejecución" & String(0.55, vbTab) & " por el solo mérito del vencimiento del plazo pactado, o de", "F1", 11, hjustify
'oDoc.WTextBox 330, 45, 30, 520, "sus renovaciones o prórrogas de ser el caso.", "F1", 11, hjustify
''-------------------------------------------------------------------------------------------
'oDoc.WTextBox 350, 45, 30, 540, "De acuerdo" & String(0.55, vbTab) & "a" & String(0.55, vbTab) & " lo dispuesto en el numeral 11) del artículo 132° de la Ley General del Sistema Financiero y del Sistema de", "F1", 11, hjustify
'oDoc.WTextBox 350, 533, 30, 540, "Seguros ", "F1", 11, hjustify
'oDoc.WTextBox 360, 45, 30, 520, "y Orgánica " & String(0.5, vbTab) & "de" & String(0.5, vbTab) & " la" & String(0.55, vbTab) & "Superintendencia" & String(0.55, vbTab) & "de" & String(0.55, vbTab) & " Banca y Seguros, autorizo(amos) a la", "F1", 11, hjustify
'oDoc.WTextBox 361, 355, 30, 480, "LA CAJA", "F2", 10, hjustify
'oDoc.WTextBox 360, 392, 30, 480, "para" & String(0.55, vbTab) & " que compense entre mis acreencias y ", "F1", 11, hjustify
'oDoc.WTextBox 370, 45, 30, 540, "activos (cuentas, valores, depósitos en general, entre otros) que" & String(0.55, vbTab) & " mantenga en su poder, hasta por el importe" & String(0.55, vbTab) & " de éste pagaré más", "F1", 11, hjustify
'oDoc.WTextBox 380, 45, 30, 540, "los intereses compensatorios, moratorios, gastos y cualquier otro concepto antes detallado en el presente título valor.", "F1", 11, hjustify
''-------------------------------------------------------------------------------------------
'oDoc.WTextBox 400, 45, 30, 530, "De" & String(0.55, vbTab) & " conformidad" & String(0.55, vbTab) & " con" & String(0.55, vbTab) & " el" & String(0.55, vbTab) & " artículo" & String(0.55, vbTab) & " 1233°" & String(0.55, vbTab) & " del" & String(0.55, vbTab) & " Código" & String(0.55, vbTab) & " Civil, acepto(amos)" & String(0.55, vbTab) & " que" & String(0.55, vbTab) & " en" & String(0.55, vbTab) & " la" & String(0.55, vbTab) & " eventualidad" & String(0.55, vbTab) & " de" & String(0.55, vbTab) & " que" & String(0.55, vbTab) & " el" & String(0.55, vbTab) & " presente" & String(0.55, vbTab) & " título se ", "F1", 11, hjustify
'oDoc.WTextBox 410, 562, 30, 530, "o", "F1", 11, hjustify
'oDoc.WTextBox 420, 45, 30, 540, "destrucción" & String(0.55, vbTab) & " parcial, deterioro" & String(0.55, vbTab) & " total, extravío" & String(0.55, vbTab) & " y sustracción, se aplicará lo dispuesto en los artículos 101° al 107° de la Ley No.27287, en lo que resultase pertinente.", "F1", 11, hjustify
'oDoc.WTextBox 410, 45, 30, 525, "perjudicara" & String(0.55, vbTab) & "por" & String(0.55, vbTab) & "cualquier" & String(0.55, vbTab) & "causa, tal" & String(1, vbTab) & "hecho" & String(1, vbTab) & "no extinguirá la obligación primitiva" & String(0.51, vbTab) & "u original. Asimismo, en caso de deterioro notable", "F1", 11, hjustify
''-------------------------------------------------------------------------------------------
'oDoc.WTextBox 450, 45, 30, 540, "Me(nos)" & String(0.55, vbTab) & " someto(emos) expresamente" & String(0.55, vbTab) & " a" & String(0.55, vbTab) & " la" & String(0.55, vbTab) & " competencia" & String(0.55, vbTab) & " y" & String(0.55, vbTab) & " tribunales" & String(0.55, vbTab) & "de" & String(0.55, vbTab) & "esta ciudad, en" & String(0.55, vbTab) & " cuyo" & String(0.55, vbTab) & " efecto" & String(0.55, vbTab) & " renuncio/renunciamos" & String(0.55, vbTab) & "al ", "F1", 11, hjustify
'oDoc.WTextBox 460, 45, 30, 540, "fuero de mi/nuestro domicilio. Señalo(amos) como domicilio aquel" & String(0.55, vbTab) & " indicado" & String(0.55, vbTab) & " en" & String(0.55, vbTab) & " este pagaré, a donde se efectuarán las diligencias", "F1", 11, hjustify
'oDoc.WTextBox 470, 45, 30, 540, "notariales, judiciales y demás que fuesen necesarias para lo que", "F1", 11, hjustify
'oDoc.WTextBox 471, 306, 30, 480, "LA CAJA", "F2", 10, hjustify
'oDoc.WTextBox 470, 342, 30, 520, "considere pertinente. Cualquier cambio de domicilio que", "F1", 11, hjustify
'oDoc.WTextBox 480, 45, 30, 540, "haga(mos), para su validez, lo haré(mos) mediante carta notarial y conforme a lo dispuesto en el artículo 40° del Código Civil.", "F1", 11, hjustify
''-------------------------------------------------------------------------------------------
'oDoc.WTextBox 500, 45, 30, 530, "Declaro(amos)" & String(2, vbTab) & " estar" & String(2, vbTab) & " plenamente" & String(2, vbTab) & " facultado(s)" & String(2, vbTab) & " para" & String(2, vbTab) & " suscribir" & String(2, vbTab) & " y" & String(2, vbTab) & " emitir" & String(2, vbTab) & "  el" & String(1, vbTab) & " presente" & String(1, vbTab) & " Pagaré, asumiendo", "F1", 11, hjustify
'oDoc.WTextBox 500, 492, 30, 480, "en" & String(1, vbTab) & " caso" & String(1, vbTab) & " contrario", "F1", 11, hjustify
'oDoc.WTextBox 510, 45, 30, 540, "responsabilidad civil y/o penal a que hubiera lugar. Se deja constancia que la información proporcionada por el(los) emitente(s) en", "F1", 11, hjustify
'oDoc.WTextBox 520, 45, 30, 540, "el presente documento, tiene" & String(0.4, vbTab) & " el" & String(0.54, vbTab) & " carácter de declaración jurada, de acuerdo con el artículo 179° de la Ley No. 26702 - Ley General del Sistema Financiero y del Sistema de Seguros y Orgánica de la Superintendencia de Banca y Seguros.", "F1", 11, hjustify
'oDoc.WTextBox 550, 45, 30, 520, "Suscribimos el presente en señal de conformidad.", "F1", 11, hjustify
''-------------------------------------------------------------------------------------------
'oDoc.WTextBox 570, 45, 15, 520, "Nombres y Apellidos/Razón Social:", "F2", 11
'Dim h As Integer
'h = 140
'
'
'If Not (RRelaCred.EOF And RRelaCred.BOF) Then
'    Do Until RRelaCred.EOF
'        If RRelaCred!nConsValor = gColRelPersTitular And (RRelaCred!nPersPersoneria = 2 Or RRelaCred!nPersPersoneria = 3 Or RRelaCred!nPersPersoneria = 4 Or RRelaCred!nPersPersoneria = 5 Or RRelaCred!nPersPersoneria = 6 Or RRelaCred!nPersPersoneria = 7 Or RRelaCred!nPersPersoneria = 8 Or RRelaCred!nPersPersoneria = 9 Or RRelaCred!nPersPersoneria = 10) Then
'
'            oDoc.WTextBox 460 + h, 45, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
'            oDoc.WTextBox 470 + h, 45, 20, 250, "Apoderado:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
'            oDoc.WTextBox 495 + h, 45, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
'            oDoc.WTextBox 513 + h, 45, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
'            oDoc.WTextBox 435 + h, 45, 35, 205, QuitarCaracter(PstaNombre(RRelaCred!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
'            oDoc.WTextBox 460 + h, 90, 15, 205, IIf(IsNull(RRelaCred!DNI), RRelaCred!Ruc, RRelaCred!DNI), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
'            oDoc.WTextBox 513 + h, 95, 35, 205, RRelaCred!cPersDireccDomicilio, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
'
'            oDoc.WTextBox 575 + h, 45, 15, 255, "Firma:_________________________", "F1", 11
'            oDoc.WTextBox 545 + h, 220, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
'            nTitu = 1
'
'            ElseIf RRelaCred!nConsValor = gColRelPersTitular Then
'            oDoc.WTextBox 460 + h, 45, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3 '
'            oDoc.WTextBox 475 + h, 45, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3 '
'            oDoc.WTextBox 435 + h, 45, 35, 205, QuitarCaracter(PstaNombre(RRelaCred!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
'            oDoc.WTextBox 460 + h, 90, 15, 205, IIf(IsNull(RRelaCred!DNI), RRelaCred!Ruc, RRelaCred!DNI), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
'            oDoc.WTextBox 475 + h, 95, 35, 205, RRelaCred!cPersDireccDomicilio, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
'            oDoc.WTextBox 570 + h, 45, 15, 255, "Firma:_________________________", "F1", 11
'            oDoc.WTextBox 540 + h, 220, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
'            nTitu = 1
'
'            nCode = 1
'
'            ElseIf (RRelaCred!nConsValor = gColRelPersConyugue Or RRelaCred!nConsValor = gColRelPersCodeudor) And (RRelaCred!nPersPersoneria = 2 Or RRelaCred!nPersPersoneria = 3 Or RRelaCred!nPersPersoneria = 4 Or RRelaCred!nPersPersoneria = 5 Or RRelaCred!nPersPersoneria = 6 Or RRelaCred!nPersPersoneria = 7 Or RRelaCred!nPersPersoneria = 8 Or RRelaCred!nPersPersoneria = 9 Or RRelaCred!nPersPersoneria = 10) Then
'            sPersCodR = RRelaCred!cPersCod
'                Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
'                If Not (RrelGar.BOF And RrelGar.EOF) Then
'                    While Not RrelGar.EOF
'                       oDoc.WTextBox 475 + h, 330, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
'                       oDoc.WTextBox 485 + h, 375, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
'                       RrelGar.MoveNext
'                    Wend
'                Else
'
'                End If
'            oDoc.WTextBox 570, 330, 15, 520, "Nombres y Apellidos/Razón Social:", "F2", 11
'            oDoc.WTextBox 435 + h, 330, 35, 205, QuitarCaracter(PstaNombre(RRelaCred!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
'            oDoc.WTextBox 460 + h, 375, 15, 205, IIf(IsNull(RRelaCred!DNI), RRelaCred!Ruc, RRelaCred!DNI), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
'
'            oDoc.WTextBox 513 + h, 380, 35, 205, RRelaCred!cPersDireccDomicilio, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
'            oDoc.WTextBox 460 + h, 330, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
'            oDoc.WTextBox 470 + h, 330, 20, 250, "Apoderado:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
'            oDoc.WTextBox 495 + h, 330, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
'            oDoc.WTextBox 513 + h, 330, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
'
'            oDoc.WTextBox 575 + h, 330, 15, 255, "Firma:_________________________", "F1", 11
'            oDoc.WTextBox 545 + h, 490, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
'
'            nCode = 1
'
'            ElseIf (RRelaCred!nConsValor = gColRelPersConyugue Or RRelaCred!nConsValor = gColRelPersCodeudor) Then
'            sPersCodR = RRelaCred!cPersCod
'                Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
'                If Not (RrelGar.BOF And RrelGar.EOF) Then
'                    While Not RrelGar.EOF
'                       oDoc.WTextBox 475 + h, 330, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
'                       oDoc.WTextBox 485 + h, 375, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
'                       RrelGar.MoveNext
'                    Wend
'                Else
'
'                End If
'            oDoc.WTextBox 570, 330, 15, 520, "Nombres y Apellidos/Razón Social:", "F2", 11
'            oDoc.WTextBox 435 + h, 330, 35, 205, QuitarCaracter(PstaNombre(RRelaCred!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
'            oDoc.WTextBox 460 + h, 375, 15, 205, IIf(IsNull(RRelaCred!DNI), RRelaCred!Ruc, RRelaCred!DNI), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
'            oDoc.WTextBox 475 + h, 380, 35, 205, RRelaCred!cPersDireccDomicilio, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
'
'
'            oDoc.WTextBox 460 + h, 330, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
''
'            oDoc.WTextBox 475 + h, 330, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
'
'            oDoc.WTextBox 570 + h, 330, 15, 255, "Firma:_________________________", "F1", 11
'            oDoc.WTextBox 540 + h, 490, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
'
'            nCode = 1
''
'            ElseIf RRelaCred!nConsValor = gColRelPersRepresentante Then
'            oDoc.WTextBox 473 + h, 45, 35, 205, QuitarCaracter(PstaNombre(RRelaCred!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
'            oDoc.WTextBox 485 + h, 90, 35, 250, IIf(IsNull(RRelaCred!DNI), RRelaCred!Ruc, RRelaCred!DNI), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
'
'            nCode = 1
'            Exit Do
'        End If
'        RRelaCred.MoveNext
'    Loop
'End If
'
'ImprimeFianza psCtaCod, lsCiudad, oDoc, RsGarantes, RelaGar
'
'oDoc.PDFClose
'oDoc.Show
'End Sub
'
Private Function ImprimeFianza(ByVal psCtaCod As String, ByVal lsCiudad As String, ByVal oDoc, ByVal RsGarantes, ByVal RelaGar)

Dim sPersCodR As String
Dim RrelGar As ADODB.Recordset
Dim h As Integer
h = 140
oDoc.NewPage A4_Vertical
oDoc.WImage 50, 494, 35, 73, "Logo"
oDoc.WTextBox 50, 45, 15, 500, "FIANZA SOLIDARIA", "F2", 11
'---------------------------------------------------------------------------------------
oDoc.WTextBox 69, 45, 15, 545, "Me/Nos constituyo/constituimos en fiador/es solidario/s del(os) emitente(s) de este Pagaré, en forma irrevocable, incondicionada, ilimitada e indefinida, a favor de la", "F1", 11, hjustify
oDoc.WTextBox 81, 185, 15, 520, "CAJA" & String(0.55, vbTab) & " MUNICIPAL" & String(0.55, vbTab) & " DE" & String(0.55, vbTab) & " AHORRO" & String(0.55, vbTab) & " Y CRÉDITO DE MAYNAS S.A.", "F2", 10, hjustify
oDoc.WTextBox 80, 430, 20, 480, ", con R.U.C. N° 20103845328, en", "F1", 11, hjustify
oDoc.WTextBox 90, 45, 20, 550, "adelante" & String(15, vbTab) & ", renunciando" & String(2, vbTab) & " expresamente" & String(2, vbTab) & " al" & String(2, vbTab) & " beneficio" & String(2, vbTab) & " de" & String(2, vbTab) & " excusión" & String(1, vbTab) & " por" & String(1, vbTab) & " obligaciones" & String(1, vbTab) & " contraídas" & String(1, vbTab) & " en" & String(1, vbTab) & " este" & String(0.55, vbTab) & " documento" & String(0.55, vbTab) & " obligándome/obligándonos" & String(0.55, vbTab) & " al" & String(0.55, vbTab) & " pago" & String(0.55, vbTab) & " de" & String(0.55, vbTab) & " la" & String(0.55, vbTab) & " cantidad" & String(0.55, vbTab) & " adeudada, intereses" & String(0.55, vbTab) & " compensatorios" & String(2, vbTab) & " y" & String(2, vbTab) & " moratorios, así como comisiones,", "F1", 11
oDoc.WTextBox 91, 82, 30, 480, "LA CAJA", "F2", 10
oDoc.WTextBox 110, 45, 20, 560, "penalidades, seguros, gastos notariales, de cobranza judicial y extrajudicial, que se" & String(0.55, vbTab) & " pudieran devengar desde la fecha de emisión ", "F1", 11, hjustify
oDoc.WTextBox 120, 45, 20, 520, "hasta la cancelación total de la presente obligación.", "F1", 11, hjustify
'---------------------------------------------------------------------------------------
oDoc.WTextBox 140, 45, 20, 540, "De" & String(0.55, vbTab) & "acuerdo" & String(0.55, vbTab) & " a" & String(0.55, vbTab) & " lo" & String(0.55, vbTab) & " dispuesto" & String(0.55, vbTab) & " en" & String(0.55, vbTab) & " el" & String(0.55, vbTab) & " numeral" & String(0.55, vbTab) & " 11)" & String(0.55, vbTab) & " del" & String(0.55, vbTab) & " articulo" & String(0.55, vbTab) & " 132°" & String(0.55, vbTab) & " de" & String(0.55, vbTab) & " la" & String(0.55, vbTab) & " Ley" & String(0.55, vbTab) & " No. 26702 - Ley General del" & String(1, vbTab) & " Sistema" & String(0.55, vbTab) & " Financiero", "F1", 11, hjustify
oDoc.WTextBox 150, 45, 20, 520, " y del Sistema de Seguros y Orgánica de la Superintendencia de Banca y Seguros, autorizo(amos) a", "F1", 11, hjustify
oDoc.WTextBox 151, 448, 30, 480, "LA CAJA", "F2", 10, hjustify
oDoc.WTextBox 150, 486, 20, 480, "para que compense", "F1", 11, hjustify
oDoc.WTextBox 160, 45, 30, 535, " entre" & String(0.55, vbTab) & " mis/nuestras" & String(0.55, vbTab) & " acreencias" & String(0.5, vbTab) & " y" & String(0.5, vbTab) & " activos (cuentas, valores, depósitos en general, entre otros) que" & String(0.55, vbTab) & " mantenga" & String(0.55, vbTab) & " en" & String(0.55, vbTab) & " su poder,", "F1", 11, hjustify
oDoc.WTextBox 160, 544, 30, 535, "hasta", "F1", 11, hjustify
oDoc.WTextBox 170, 45, 30, 540, "por" & String(0.55, vbTab) & " el" & String(0.55, vbTab) & " importe" & String(0.55, vbTab) & " adeudado" & String(0.55, vbTab) & " de" & String(0.5, vbTab) & " este pagaré más los intereses compensatorios, moratorios, gastos, y cualquier otro concepto que puedan generarse.", "F1", 11, hjustify
'---------------------------------------------------------------------------------------
oDoc.WTextBox 200, 45, 30, 530, "Asimismo, me(nos)" & String(0.55, vbTab) & " someto(emos)" & String(0.55, vbTab) & " expresamente" & String(0.55, vbTab) & " a" & String(0.55, vbTab) & " la" & String(0.55, vbTab) & " competencia" & String(0.55, vbTab) & " y" & String(2, vbTab) & " tribunales" & String(2, vbTab) & " de" & String(2, vbTab) & " esta" & String(2, vbTab) & " ciudad, en" & String(1, vbTab) & " cuyo" & String(1, vbTab) & " efecto" & String(1, vbTab) & " renuncio", "F1", 11, hjustify
oDoc.WTextBox 210, 45, 30, 540, "/renunciamos al fuero de mi/nuestro domicilio. Señalo(amos) como domicilio" & String(0.55, vbTab) & " aquel" & String(0.5, vbTab) & " indicado en este pagaré a donde se efectuarán las diligencias notariales, judiciales y demás que fuesen necesarias para lo que ", "F1", 11, hjustify
oDoc.WTextBox 221, 365, 30, 480, "LA CAJA", "F2", 10, hjustify
oDoc.WTextBox 220, 400, 30, 480, "considere pertinente.Cualquier cambio de", "F1", 11, hjustify
oDoc.WTextBox 230, 45, 30, 540, " domicilio" & String(0.55, vbTab) & " que" & String(0.55, vbTab) & " haga(mos), para" & String(0.5, vbTab) & " su" & String(0.53, vbTab) & "validez, lo" & String(0.55, vbTab) & " haré(mos), mediante" & String(0.55, vbTab) & "carta" & String(0.55, vbTab) & " notarial" & String(0.55, vbTab) & " y" & String(0.55, vbTab) & " conforme" & String(0.55, vbTab) & " a lo dispuesto en el artículo 40° ", "F1", 11, hjustify
oDoc.WTextBox 230, 555, 30, 34, "del", "F1", 11, hjustify
oDoc.WTextBox 240, 45, 30, 520, "Código Civil.", "F1", 11, hjustify
'---------------------------------------------------------------------------------------
oDoc.WTextBox 260, 45, 30, 540, "Declaro(amos)" & String(0.55, vbTab) & " estar plenamente facultado(s) para afianzar el presente Pagaré, asumiendo en caso contrario la responsabilidad civil y/o penal que hubiere lugar. Se" & String(0.5, vbTab) & " deja" & String(0.55, vbTab) & "constancia que la información proporciona por el(los) fiador(es) en el presente documento, tiene el carácter de declaración jurada, de acuerdo con el artículo 179° de la Ley N° 26702.", "F1", 11, hjustify
oDoc.WTextBox 305, 45, 30, 480, "Suscribimos el presente en señal de conformidad.", "F1", 11, hjustify


If lsCiudad = "MOYOBAMBA " Or lsCiudad = "CAJAMARCA " Or lsCiudad = "YURIMAGUAS " Or lsCiudad = "TINGO MARIA " Then
oDoc.WTextBox 330, 45, 30, 480, "Lugar y fecha de emisión:" & String(23, vbTab) & " ,___de_____________________de _____", "F1", 11, hjustify
oDoc.WTextBox 332, 96, 15, 160, lsCiudad, "F2", 10, vMiddle, vbBlack, 1, vbBlack
ElseIf lsCiudad = "PUERTO CALLAO " Or lsCiudad = "CERRO DE PASCO " Or lsCiudad = "TOCACHE NUEVO " Then
oDoc.WTextBox 330, 45, 30, 480, "Lugar y fecha de emisión:" & String(30, vbTab) & " ,____de_____________________de _____", "F1", 11, hjustify
oDoc.WTextBox 332, 105, 15, 160, lsCiudad, "F2", 10, vMiddle, vbBlack, 1, vbBlack
Else: oDoc.WTextBox 330, 45, 30, 480, "Lugar y fecha de emisión:" & String(20, vbTab) & " ,___de_____________________de _____", "F1", 11, hjustify
oDoc.WTextBox 332, 90, 15, 160, lsCiudad, "F2", 10, vMiddle, vbBlack, 1, vbBlack
End If
           
Dim nGaran As Integer
nGaran = 0

If Not (RsGarantes.EOF And RsGarantes.BOF) Then
        While Not RsGarantes.EOF
        '##################################### 1 Garante ##############################################
            If nGaran = 0 And (RsGarantes!nPersPersoneria = 2 Or RsGarantes!nPersPersoneria = 3 Or RsGarantes!nPersPersoneria = 4 Or RsGarantes!nPersPersoneria = 5 Or RsGarantes!nPersPersoneria = 6 Or RsGarantes!nPersPersoneria = 7 Or RsGarantes!nPersPersoneria = 8 Or RsGarantes!nPersPersoneria = 9 Or RsGarantes!nPersPersoneria = 10) Then
                sPersCodR = RsGarantes!cperscod
                Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR) '----Recupera el Representante
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 285 + h, 45, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 295 + h, 90, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                End If
                
             oDoc.WTextBox 245 + h, 45, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 270 + h, 90, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 320 + h, 95, 35, 205, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                '1
             oDoc.WTextBox 360, 45, 15, 450, "FIADOR SOLIDARIO", "F2", 11
             oDoc.WTextBox 380, 45, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11 '
             oDoc.WTextBox 270 + h, 45, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3 '
             oDoc.WTextBox 280 + h, 45, 20, 60, "Apoderado:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3 '
             oDoc.WTextBox 305 + h, 45, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3 '
             oDoc.WTextBox 320 + h, 45, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
'
             oDoc.WTextBox 365 + h, 45, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 345 + h, 220, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
             
            ElseIf nGaran = 0 Then
            sPersCodR = RsGarantes!cperscod
                Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 285 + h, 45, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 295 + h, 90, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                   
                End If
                
             oDoc.WTextBox 245 + h, 45, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 270 + h, 90, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 285 + h, 95, 35, 205, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                '1
             oDoc.WTextBox 360, 45, 15, 450, "FIADOR SOLIDARIO", "F2", 11
             oDoc.WTextBox 380, 45, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11

             oDoc.WTextBox 270 + h, 45, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 285 + h, 45, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 360 + h, 45, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 340 + h, 220, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
              '##################################### 2 Garante ##############################################
            ElseIf nGaran = 1 And (RsGarantes!nPersPersoneria = 2 Or RsGarantes!nPersPersoneria = 3 Or RsGarantes!nPersPersoneria = 4 Or RsGarantes!nPersPersoneria = 5 Or RsGarantes!nPersPersoneria = 6 Or RsGarantes!nPersPersoneria = 7 Or RsGarantes!nPersPersoneria = 8 Or RsGarantes!nPersPersoneria = 9 Or RsGarantes!nPersPersoneria = 10) Then
            sPersCodR = RsGarantes!cperscod
                Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 285 + h, 300, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 295 + h, 345, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                    
                End If
                oDoc.WTextBox 245 + h, 300, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 270 + h, 345, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 320 + h, 350, 35, 205, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                 '2
             oDoc.WTextBox 360, 300, 15, 450, "FIADOR SOLIDARIO", "F2", 11
             oDoc.WTextBox 380, 300, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 270 + h, 300, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 280 + h, 300, 20, 60, "Apoderado:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 305 + h, 300, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 320 + h, 300, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 365 + h, 300, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 345 + h, 490, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
             
            ElseIf nGaran = 1 Then
            sPersCodR = RsGarantes!cperscod
                Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 285 + h, 300, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 295 + h, 345, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                    
                End If
                oDoc.WTextBox 245 + h, 300, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 270 + h, 345, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 285 + h, 350, 35, 205, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                 '2
             oDoc.WTextBox 360, 300, 15, 450, "FIADOR SOLIDARIO", "F2", 11
             oDoc.WTextBox 380, 300, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11

             oDoc.WTextBox 270 + h, 300, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3

             oDoc.WTextBox 285 + h, 300, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3

             oDoc.WTextBox 360 + h, 300, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 340 + h, 490, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
          '##################################### 3 Garante ##############################################
            ElseIf nGaran = 2 And (RsGarantes!nPersPersoneria = 2 Or RsGarantes!nPersPersoneria = 3 Or RsGarantes!nPersPersoneria = 4 Or RsGarantes!nPersPersoneria = 5 Or RsGarantes!nPersPersoneria = 6 Or RsGarantes!nPersPersoneria = 7 Or RsGarantes!nPersPersoneria = 8 Or RsGarantes!nPersPersoneria = 9 Or RsGarantes!nPersPersoneria = 10) Then
            
            sPersCodR = RsGarantes!cperscod
                Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 505 + h, 45, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 515 + h, 90, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                    
                End If
                oDoc.WTextBox 465 + h, 45, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 490 + h, 90, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 540 + h, 95, 35, 180, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                '2
              oDoc.WTextBox 580, 45, 15, 450, "FIADOR SOLIDARIO", "F2", 11
              oDoc.WTextBox 600, 45, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 490 + h, 45, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 500 + h, 45, 20, 60, "Apoderado:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 525 + h, 45, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 540 + h, 45, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 605 + h, 45, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 585 + h, 220, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack


            ElseIf nGaran = 2 Then
            
            sPersCodR = RsGarantes!cperscod
                Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 505 + h, 45, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 515 + h, 90, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                    
                End If
                oDoc.WTextBox 465 + h, 45, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 490 + h, 90, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 505 + h, 95, 35, 180, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                '2
              oDoc.WTextBox 580, 45, 15, 450, "FIADOR SOLIDARIO", "F2", 11
              oDoc.WTextBox 600, 45, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 490 + h, 45, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 505 + h, 45, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 605 + h, 45, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 585 + h, 220, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
           '##################################### 4 Garante ##############################################
            ElseIf nGaran = 3 And (RsGarantes!nPersPersoneria = 2 Or RsGarantes!nPersPersoneria = 3 Or RsGarantes!nPersPersoneria = 4 Or RsGarantes!nPersPersoneria = 5 Or RsGarantes!nPersPersoneria = 6 Or RsGarantes!nPersPersoneria = 7 Or RsGarantes!nPersPersoneria = 8 Or RsGarantes!nPersPersoneria = 9 Or RsGarantes!nPersPersoneria = 10) Then
            sPersCodR = RsGarantes!cperscod
            Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 505 + h, 300, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 515 + h, 345, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                    
                End If
                oDoc.WTextBox 465 + h, 300, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 490 + h, 345, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 538 + h, 350, 35, 205, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                
                '4
             oDoc.WTextBox 580, 300, 15, 450, "FIADOR SOLIDARIO", "F2", 11
             oDoc.WTextBox 600, 300, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 490 + h, 300, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 500 + h, 300, 20, 60, "Apoderado:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 525 + h, 300, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 538 + h, 300, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 605 + h, 300, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 585 + h, 490, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
             
             ElseIf nGaran = 3 Then
            sPersCodR = RsGarantes!cperscod
            Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 505 + h, 300, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 515 + h, 345, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                    
                End If
                oDoc.WTextBox 465 + h, 300, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 490 + h, 345, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 505 + h, 350, 35, 205, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                
                '4
             oDoc.WTextBox 580, 300, 15, 450, "FIADOR SOLIDARIO", "F2", 11
             oDoc.WTextBox 600, 300, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 490 + h, 300, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 505 + h, 300, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 605 + h, 300, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 585 + h, 490, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
             '##################################### 5 Garante ##############################################
            
            ElseIf nGaran = 4 And (RsGarantes!nPersPersoneria = 2 Or RsGarantes!nPersPersoneria = 3 Or RsGarantes!nPersPersoneria = 4 Or RsGarantes!nPersPersoneria = 5 Or RsGarantes!nPersPersoneria = 6 Or RsGarantes!nPersPersoneria = 7 Or RsGarantes!nPersPersoneria = 8 Or RsGarantes!nPersPersoneria = 9 Or RsGarantes!nPersPersoneria = 10) Then
            h = -170
            oDoc.NewPage A4_Vertical
            oDoc.WImage 50, 494, 35, 73, "Logo"
            
                sPersCodR = RsGarantes!cperscod
                Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 283 + h, 45, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 295 + h, 90, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                   
                End If
                            
                
             oDoc.WTextBox 245 + h, 45, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 270 + h, 90, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 320 + h, 95, 35, 205, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                '1
             oDoc.WTextBox 50, 45, 15, 450, "FIADOR SOLIDARIO", "F2", 11
             oDoc.WTextBox 70, 45, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 270 + h, 45, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 280 + h, 45, 20, 60, "Apoderado:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 305 + h, 45, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 320 + h, 45, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 380 + h, 45, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 360 + h, 220, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
            ElseIf nGaran = 4 Then
            h = -170
            oDoc.NewPage A4_Vertical
            
            sPersCodR = RsGarantes!cperscod
                Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 283 + h, 45, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 295 + h, 90, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                   
                End If
                            
                
             oDoc.WTextBox 245 + h, 45, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 270 + h, 90, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 285 + h, 95, 35, 205, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                '1
             oDoc.WTextBox 50, 45, 15, 450, "FIADOR SOLIDARIO", "F2", 11
             oDoc.WTextBox 70, 45, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 270 + h, 45, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 285 + h, 45, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 380 + h, 45, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 360 + h, 220, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
             '##################################### 6 Garante ##############################################
            ElseIf nGaran = 5 And (RsGarantes!nPersPersoneria = 2 Or RsGarantes!nPersPersoneria = 3 Or RsGarantes!nPersPersoneria = 4 Or RsGarantes!nPersPersoneria = 5 Or RsGarantes!nPersPersoneria = 6 Or RsGarantes!nPersPersoneria = 7 Or RsGarantes!nPersPersoneria = 8 Or RsGarantes!nPersPersoneria = 9 Or RsGarantes!nPersPersoneria = 10) Then
            sPersCodR = RsGarantes!cperscod
                Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 283 + h, 300, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 295 + h, 345, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                    
                End If
                oDoc.WTextBox 245 + h, 300, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 270 + h, 345, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 320 + h, 350, 35, 205, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                 '2
             oDoc.WTextBox 50, 300, 15, 450, "FIADOR SOLIDARIO", "F2", 11
             oDoc.WTextBox 70, 300, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 270 + h, 300, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 280 + h, 300, 20, 60, "Apoderado:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 305 + h, 300, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 320 + h, 300, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 380 + h, 300, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 360 + h, 490, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
             
            ElseIf nGaran = 5 Then
            
            sPersCodR = RsGarantes!cperscod
                Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 283 + h, 300, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 295 + h, 345, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                    
                End If
                oDoc.WTextBox 245 + h, 300, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 270 + h, 345, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 285 + h, 350, 35, 205, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                 '2
             oDoc.WTextBox 50, 300, 15, 450, "FIADOR SOLIDARIO", "F2", 11
             oDoc.WTextBox 70, 300, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 270 + h, 300, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 285 + h, 300, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 380 + h, 300, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 360 + h, 490, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
             '##################################### 7 Garante ##############################################
            ElseIf nGaran = 6 And (RsGarantes!nPersPersoneria = 2 Or RsGarantes!nPersPersoneria = 3 Or RsGarantes!nPersPersoneria = 4 Or RsGarantes!nPersPersoneria = 5 Or RsGarantes!nPersPersoneria = 6 Or RsGarantes!nPersPersoneria = 7 Or RsGarantes!nPersPersoneria = 8 Or RsGarantes!nPersPersoneria = 9 Or RsGarantes!nPersPersoneria = 10) Then
            
            sPersCodR = RsGarantes!cperscod
                Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 523 + h, 45, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 535 + h, 90, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                    
                End If
                oDoc.WTextBox 485 + h, 45, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 510 + h, 90, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 560 + h, 95, 35, 180, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                '2
              oDoc.WTextBox 290, 45, 15, 450, "FIADOR SOLIDARIO", "F2", 11
              oDoc.WTextBox 310, 45, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 510 + h, 45, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 520 + h, 45, 20, 60, "Apoderado:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 545 + h, 45, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 560 + h, 45, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 625 + h, 45, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 605 + h, 220, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
            ElseIf nGaran = 6 Then
            h = -170
            sPersCodR = RsGarantes!cperscod
                Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 523 + h, 45, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 535 + h, 90, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                    
                End If
                oDoc.WTextBox 485 + h, 45, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 510 + h, 90, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 525 + h, 95, 35, 180, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                '2
              oDoc.WTextBox 290, 45, 15, 450, "FIADOR SOLIDARIO", "F2", 11
              oDoc.WTextBox 310, 45, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 510 + h, 45, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 525 + h, 45, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 625 + h, 45, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 605 + h, 220, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
             '##################################### 8 Garante ##############################################
           ElseIf nGaran = 7 And (RsGarantes!nPersPersoneria = 2 Or RsGarantes!nPersPersoneria = 3 Or RsGarantes!nPersPersoneria = 4 Or RsGarantes!nPersPersoneria = 5 Or RsGarantes!nPersPersoneria = 6 Or RsGarantes!nPersPersoneria = 7 Or RsGarantes!nPersPersoneria = 8 Or RsGarantes!nPersPersoneria = 9 Or RsGarantes!nPersPersoneria = 10) Then
            sPersCodR = RsGarantes!cperscod
            Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 523 + h, 300, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 535 + h, 345, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                    
                End If
                oDoc.WTextBox 485 + h, 300, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 510 + h, 345, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 560 + h, 350, 35, 205, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                
                '8
             oDoc.WTextBox 290, 300, 15, 450, "FIADOR SOLIDARIO", "F2", 11
             oDoc.WTextBox 310, 300, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 510 + h, 300, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 520 + h, 300, 20, 60, "Apoderado:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 545 + h, 300, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 560 + h, 300, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 625 + h, 300, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 605 + h, 490, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
             
             ElseIf nGaran = 7 Then
            sPersCodR = RsGarantes!cperscod
            Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 523 + h, 300, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 535 + h, 345, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                    
                End If
                oDoc.WTextBox 485 + h, 300, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 510 + h, 345, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 525 + h, 350, 35, 205, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4

                '4
             oDoc.WTextBox 290, 300, 15, 450, "FIADOR SOLIDARIO", "F2", 11
             oDoc.WTextBox 310, 300, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 510 + h, 300, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 525 + h, 300, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 625 + h, 300, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 605 + h, 490, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
              '##################################### 9 Garante ##############################################
            ElseIf nGaran = 8 And (RsGarantes!nPersPersoneria = 2 Or RsGarantes!nPersPersoneria = 3 Or RsGarantes!nPersPersoneria = 4 Or RsGarantes!nPersPersoneria = 5 Or RsGarantes!nPersPersoneria = 6 Or RsGarantes!nPersPersoneria = 7 Or RsGarantes!nPersPersoneria = 8 Or RsGarantes!nPersPersoneria = 9 Or RsGarantes!nPersPersoneria = 10) Then
            h = 110
            sPersCodR = RsGarantes!cperscod
                Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 503 + h, 45, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 515 + h, 90, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                    
                End If
                oDoc.WTextBox 465 + h, 45, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 490 + h, 90, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 542 + h, 95, 35, 180, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                '2
              oDoc.WTextBox 550, 45, 15, 450, "FIADOR SOLIDARIO", "F2", 11
              oDoc.WTextBox 570, 45, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 490 + h, 45, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 500 + h, 45, 20, 60, "Apoderado:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 525 + h, 45, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 542 + h, 45, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 605 + h, 45, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 585 + h, 220, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack


            ElseIf nGaran = 8 Then
            
            sPersCodR = RsGarantes!cperscod
                Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 503 + h, 45, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 515 + h, 90, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                    
                End If
                oDoc.WTextBox 465 + h, 45, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 490 + h, 90, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 505 + h, 95, 35, 180, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                '2
              oDoc.WTextBox 550, 45, 15, 450, "FIADOR SOLIDARIO", "F2", 11
              oDoc.WTextBox 570, 45, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 490 + h, 45, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 505 + h, 45, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 605 + h, 45, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 585 + h, 220, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
           '##################################### 10 Garante ##############################################
            ElseIf nGaran = 9 And (RsGarantes!nPersPersoneria = 2 Or RsGarantes!nPersPersoneria = 3 Or RsGarantes!nPersPersoneria = 4 Or RsGarantes!nPersPersoneria = 5 Or RsGarantes!nPersPersoneria = 6 Or RsGarantes!nPersPersoneria = 7 Or RsGarantes!nPersPersoneria = 8 Or RsGarantes!nPersPersoneria = 9 Or RsGarantes!nPersPersoneria = 10) Then
            sPersCodR = RsGarantes!cperscod
            Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 503 + h, 300, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 515 + h, 345, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                    
                End If
                oDoc.WTextBox 465 + h, 300, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 490 + h, 345, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 538 + h, 350, 35, 205, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                
                '4
             oDoc.WTextBox 550, 300, 15, 450, "FIADOR SOLIDARIO", "F2", 11
             oDoc.WTextBox 570, 300, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 490 + h, 300, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 500 + h, 300, 20, 60, "Apoderado:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 525 + h, 300, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 538 + h, 300, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 605 + h, 300, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 585 + h, 490, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
             
             ElseIf nGaran = 9 Then
            sPersCodR = RsGarantes!cperscod
            Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 503 + h, 300, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 515 + h, 345, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                    
                End If
                oDoc.WTextBox 465 + h, 300, 35, 205, QuitarCaracter(PstaNombre(RsGarantes!cPersNombre), "-*/\._"), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 490 + h, 345, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 505 + h, 350, 35, 205, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                
                '4
             oDoc.WTextBox 550, 300, 15, 450, "FIADOR SOLIDARIO", "F2", 11
             oDoc.WTextBox 570, 300, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 490 + h, 300, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 505 + h, 300, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 605 + h, 300, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 585 + h, 490, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
            End If
            nGaran = nGaran + 1
            RsGarantes.MoveNext
        Wend
    Else
End If
'########################PTI1 20170315  ########################################
End Function
Private Function QuitarCaracter(ByVal psCadena As String, ByVal psCaracter As String) As String
Dim nPosicion As Integer
Dim nTamano As Integer
Dim i As Integer
Dim sResultado As String
Dim sTemp As String
sResultado = psCadena

nTamano = Len(psCaracter)
nPosicion = Len(psCaracter)
sTemp = psCaracter

For i = 0 To nTamano - 1
    sTemp = Mid(sTemp, i + 1, 1)
    sResultado = Replace(sResultado, sTemp, "")
    sTemp = psCaracter
Next i

QuitarCaracter = sResultado
End Function

''*** PEAC 20170621 - SE ENVIO A gCredReportes
''WIOR FIN ******************************************************************************************************
''Agregado por PASI20131127 segun TI-ERS136-2013
'Public Function VerificarExisteDesembolsoBcoNac(ByVal psCtaCod As String, ByRef sFecDes As String, ByVal pnOpcion As Integer) As Boolean
'    Dim oDCred As COMNCredito.NCOMCredito
'    Dim oCredDoc As COMDCredito.DCOMCredDoc
'    Dim oDCred2 As COMDCredito.DCOMCredito
'    Dim bValor As Boolean
'    Dim R As ADODB.Recordset
'    Set oDCred = New COMNCredito.NCOMCredito
'
'    bValor = oDCred.VerificarExisteDesembolsoBcoNac(psCtaCod)
'    If bValor = True Then
'        If (MsgBox("El Crédito no ha sido Desembolsado por lo que no cuenta con fecha para la generación del documento; desea agregar manualmente?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmación") = vbYes) Then
'
'            If pnOpcion = 1 Then
'                Set oDCred2 = New COMDCredito.DCOMCredito
'                Set R = oDCred2.RecuperaDatosComunes(psCtaCod)
'                sFecDes = frmIngFechaGenDoc.Inicio(R!dVigencia)
'            End If
'
'            If pnOpcion = 2 Then
'                Set oCredDoc = New COMDCredito.DCOMCredDoc
'                Set R = oCredDoc.RecuperaDatosDocPlanPagos(psCtaCod)
'                sFecDes = frmIngFechaGenDoc.Inicio(R!dFecVig)
'            End If
'
'            Set R = Nothing
'
'            If (sFecDes = Empty) Then
'                VerificarExisteDesembolsoBcoNac = False
'            Else
'                VerificarExisteDesembolsoBcoNac = True
'            End If
'        Else
'        VerificarExisteDesembolsoBcoNac = False '********Con esto ocurre el proceso normal
'        End If
'    Else
'        VerificarExisteDesembolsoBcoNac = False '********Con esto ocurre el proceso normal
'    End If
'End Function
'************************************************************


'RECO20140305 ERS174-2013*************************************
Private Sub ImprimeHojaAprobacionCred(ByRef pR As ADODB.Recordset, ByRef pRB As ADODB.Recordset, ByRef pRCliRela As ADODB.Recordset, ByRef pRRelaBcos As ADODB.Recordset, _
                                      ByRef pRDatFinan As ADODB.Recordset, ByRef pRResGarTitAva As ADODB.Recordset, ByRef pRGarantCred As ADODB.Recordset, _
                                      ByVal psCodCta, ByVal prsCredEval As ADODB.Recordset, ByVal prsCredAmp As ADODB.Recordset, ByVal prsCalfSBSRela As ADODB.Recordset, _
                                      ByVal prsExoAutCred As ADODB.Recordset, ByVal prsRiesgoUnico As ADODB.Recordset, ByVal prsCredGarant As ADODB.Recordset, _
                                      ByVal prsCredResulNivApr As ADODB.Recordset, ByVal pRNivApr As ADODB.Recordset, ByVal psRiesgo As String, ByVal psHabilitaNiveles As String, _
                                      ByVal prsOpRiesgo As ADODB.Recordset, ByVal prsComentAnalis As ADODB.Recordset, _
                                      ByVal pnTpoCambio As Double, ByVal prsSobreEnd As ADODB.Recordset, ByVal prsSobreEndCodigos As ADODB.Recordset, _
                                      ByVal pDRRatiosF As ADODB.Recordset, ByVal prsAutorizaciones As ADODB.Recordset, ByVal prsPersVinc As ADODB.Recordset) 'RECO20140804 Se añadió el parametro pnTpoCambio
                                      'WIOR 20160621 AGREGO ByVal prsSobreEnd As ADODB.Recordset, ByVal prsSobreEndCodigos As ADODB.Recordset
                                      'RECO20160730 SE AGREGO RATIOS DE FORMATO EVALUACION
                                      'FRHU 20160811 Anexo002 ERS002-2016: Se agrego ByVal prsAutorizaciones As ADODB.Recordset
                                      'APRI 20170719 AGREGO prsPersVinc TI-ERS-025 2017
    Dim oFun As New COMFunciones.FCOMImpresion
    Dim lnliquidez As Double, lnCapacidadPago As Double, lnExcedente As Double
    Dim lnPatriEmpre As Double, lnPatrimonio As Double, lnIngresoNeto As Double
    Dim lnRentabPatrimonial As Double, lnEndeudamiento As Double
    Dim lnInventario As Double, lnCuota As Double, lnRemNeta As Double, lnEgresos As Double
    'ALPA20160714****************************
    Dim lnCapacidadPago2 As Double
    '****************************************
    Dim lnMontoRiesgoUnico As Double
    Dim lnMontoExpEstCred As Double
    
    'Dim oCliPre As New COMNCredito.NCOMCredito 'COMENTADO POR ARLO 20170722
    Dim oDCredExoAut As New COMDCredito.DCOMNivelAprobacion
    Dim lrsRespNvlExo As ADODB.Recordset
    
    'CTI3 (ferimoro)
    Dim oDpExposicion As COMDCredito.DCOMCredito 'CTI3 (ferimoro) ERS062-2018
    Dim expo As ADODB.Recordset                 'CTI3 (ferimoro) ERS062-2018
    Dim nRunico As Integer
    '**************
    
    Set oDCredExoAut = New COMDCredito.DCOMNivelAprobacion
    Dim bValidarCliPre As Boolean
    
    On Error GoTo ErrorImprimirPDF
    Dim oDoc  As cPDF
    Dim nTipo As Integer
    Set oDoc = New cPDF
    
    
    Dim i As Integer
    Dim a As Integer
    Dim nPosicion As Integer
    
    
    Dim lrsIngGas As ADODB.Recordset
    Dim rsConDPer As COMDPersona.DCOMPersonas
    Dim oCrDoc As New COMDCredito.DCOMCredDoc
    Dim dFecFteIng As Date
    
    Dim lrDatosRatiosFI As ADODB.Recordset
    
    Dim nIngNeto As Double
    Dim nGasFamiliar  As Double
    Dim nExpoCredAct As Double
    
    Dim nValorR1 As Integer
    Dim nValorR2 As Integer
    
    Dim nClePref As Integer
    
    If oCrDoc.ObtenerFechaProNumFuente(pR!cNumeroFuente, gdFecSis).RecordCount > 0 Then
        dFecFteIng = Format(oCrDoc.ObtenerFechaProNumFuente(pR!cNumeroFuente, gdFecSis)!valor, "dd/mm/yyyy")
    End If
    
    Set lrDatosRatiosFI = New ADODB.Recordset
    Set lrDatosRatiosFI = oCrDoc.ObtieneRatiosFI(pR!cNumeroFuente, psCodCta)
    
    lnliquidez = 0
    lnEndeudamiento = 0
    lnRentabPatrimonial = 0
    lnCapacidadPago = 0
    lnPatrimonio = 0
    lnInventario = 0
    lnExcedente = 0
    lnCuota = 0
    lnRemNeta = 0
    lnEgresos = 0
    
    nValorR1 = 0 'JOEP20170627
    nValorR2 = 0 'JOEP20170627
    
    If Not (lrDatosRatiosFI.EOF And lrDatosRatiosFI.BOF) Then
        Dim nIndexRatios As Integer
        For nIndexRatios = 0 To lrDatosRatiosFI.RecordCount - 1
            If lrDatosRatiosFI!nValor = 1 Then
                lnInventario = lrDatosRatiosFI!nMonto
            ElseIf lrDatosRatiosFI!nValor = 2 Then
                lnExcedente = lrDatosRatiosFI!nMonto
            ElseIf lrDatosRatiosFI!nValor = 3 Then
                lnCuota = lrDatosRatiosFI!nMonto
            ElseIf lrDatosRatiosFI!nValor = 4 Then
                lnPatrimonio = lrDatosRatiosFI!nMonto
            ElseIf lrDatosRatiosFI!nValor = 5 Then
                lnEndeudamiento = lrDatosRatiosFI!nMonto
            ElseIf lrDatosRatiosFI!nValor = 6 Then
                lnliquidez = lrDatosRatiosFI!nMonto
            ElseIf lrDatosRatiosFI!nValor = 9 Then
                lnRemNeta = lrDatosRatiosFI!nMonto
            ElseIf lrDatosRatiosFI!nValor = 10 Then
                lnEgresos = lrDatosRatiosFI!nMonto
            'ALPA20160714*****************************************************************************
            ElseIf lrDatosRatiosFI!nValor = 11 Then
                lnCapacidadPago2 = IIf(IsNull(lrDatosRatiosFI!nMonto), 0, lrDatosRatiosFI!nMonto)
            '*****************************************************************************************
            End If
            
            lrDatosRatiosFI.MoveNext
        Next
        'ALPA20160714******************************************************************
        'lnCapacidadPago = lnCuota / IIf(lnExcedente = 0, 1, lnExcedente)
        If lnCapacidadPago2 = 0 Then
            lnCapacidadPago = lnCuota / IIf(lnExcedente = 0, 1, lnExcedente)
        Else
            lnCapacidadPago = lnCapacidadPago2
        End If
        '******************************************************************************
        lnRentabPatrimonial = IIf(lnExcedente = 0, 1, lnExcedente) / IIf(lnPatrimonio = 0, 1, lnPatrimonio)
        lnEndeudamiento = lnEndeudamiento / IIf(lnPatrimonio = 0, 1, lnPatrimonio)
    End If
    Set rsConDPer = New COMDPersona.DCOMPersonas
    Set lrsIngGas = rsConDPer.ObtenerDatosHojEvaluaci(pR!cNumeroFuente, dFecFteIng)



    oDoc.Author = gsCodUser
    oDoc.Creator = "SICMACT - Negocio"
    oDoc.Producer = "Caja Municipal de Ahorros y Crédito de Maynas S.A."
    oDoc.Subject = "Hoja de Aprobación de Créditos Nº " & gsCodUser
    oDoc.Title = "Hoja de Aprobación de Créditos Nº " & gsCodUser
    If Not oDoc.PDFCreate(App.Path & "\Spooler\" & IIf(nTipo = 1, "Previo", "") & psCodCta & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
        Exit Sub
    End If

    'FUENTES
    Dim nFTabla As Integer
    Dim nFTablaCabecera As Integer
    Dim lnFontSizeBody As Integer
    
    oDoc.Fonts.Add "F1", "Times New Roman", TrueType, Normal, WinAnsiEncoding 'FRHU20131126
    oDoc.Fonts.Add "F2", "Times New Roman", TrueType, Bold, WinAnsiEncoding 'FRHU20131126
    oDoc.Fonts.Add "F3", "Arial Narrow", TrueType, Normal, WinAnsiEncoding 'RECO20140227
    oDoc.Fonts.Add "F4", "Arial Narrow", TrueType, Bold, WinAnsiEncoding 'RECO20140227
    nFTablaCabecera = 7
    nFTabla = 7
    lnFontSizeBody = 7
    'FIN FUENTES

    oDoc.NewPage A4_Vertical
    lnMontoRiesgoUnico = 0#



        'Set oCliPre = New COMNCredito.NCOMCredito           'COMENTADO POR ARLO 20170722
        'bValidarCliPre = oCliPre.ValidarClientePreferencial(pR!cPersCod) 'COMENTADO POR ARLO 20170722
        bValidarCliPre = False 'ARLO 20170722

    If Not (prsRiesgoUnico.EOF And prsRiesgoUnico.BOF) Then
        lnMontoRiesgoUnico = prsRiesgoUnico!nMonto
    End If

          If Not (lrsIngGas.BOF And lrsIngGas.EOF) Then

            Dim j As Integer

            For j = 1 To lrsIngGas.RecordCount
                If (lrsIngGas!cCodHojEval = 300102) Then
                    nIngNeto = Format(lrsIngGas!nUnico, gcFormView)
                ElseIf (lrsIngGas!cCodHojEval = 300201) Then
                    nGasFamiliar = Format(lrsIngGas!nUnico, gcFormView)
                ElseIf (lrsIngGas!cCodHojEval = 400205) Then
                    nExpoCredAct = Format(lrsIngGas!nUnico, gcFormView)
                End If
                lrsIngGas.MoveNext
            Next
         End If
    
    'oDoc.WTextBox 40, 70, 10, 450, "HOJA DE APROBACIÓN", "F3", 11, hCenter, , vbBlack

    'RECO20141011 ***********************************
    If bValidarCliPre = True Then
        nClePref = 2 'cti3
        oDoc.WTextBox 40, 70, 10, 450, "HOJA DE APROBACIÓN - CLIENTE PREFERENCIAL", "F3", 11, hCenter, , vbBlack
    Else
        nClePref = 1
        oDoc.WTextBox 40, 70, 10, 450, "HOJA DE APROBACIÓN", "F3", 11, hCenter, , vbBlack
    End If
    'RECO END ***************************************

    oDoc.WTextBox 60, 55, 10, 200, "APROBACION DE CREDITOS", "F4", 9, hLeft
    oDoc.WTextBox 60, 354, 10, 200, "FECHA APROBACION:....................", "F4", 9, hRight

    'SECCION Nº 1
        oDoc.WTextBox 80, 55, 77, 510, "", "F3", 7, hLeft, , , 1, vbBlack
    'IZQUIERDA
        'ETIQUETAS
            oDoc.WTextBox 80, 57, 12, 450, "Cliente:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 88, 57, 12, 450, "DNI/RUC:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 96, 57, 12, 450, "Tipo de Credito:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 104, 57, 12, 450, "Producto de Crédito:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 112, 57, 12, 450, "Exposición Anterior Máxima:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 120, 57, 12, 450, "Exposición con este Crédito:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 128, 57, 12, 450, "Exposición Riesgo Único:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 136, 57, 12, 450, "Fecha de Solicitud:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 144, 57, 12, 450, "Campaña", "F3", lnFontSizeBody, hLeft, , , , , , 2
            
        'DATOS
            oDoc.WTextBox 80, 140, 12, 450, UCase(pR!Prestatario), "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 88, 140, 12, 450, pR!DniRuc & "/" & pR!Ruc, "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 96, 140, 12, 450, pR!Tipo_Cred, "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 104, 140, 12, 450, pR!Tipo_Prod, "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 112, 140, 12, 450, Format(pR!ExpoAntMax, gcFormView), "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 120, 140, 12, 450, Format(pR!nMontoExpCredito, gcFormView), "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 128, 140, 12, 450, Format(lnMontoRiesgoUnico, gcFormView), "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 136, 140, 12, 450, pR!Fec_Soli, "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 144, 140, 12, 450, pR!cCampania, "F3", lnFontSizeBody, hLeft, , , , , , 2
            lnMontoExpEstCred = pR!nMontoExpCredito

'****************************************************************
'Variables Capturadas para ser usadas
'pR!DniRuc 'dni
'pR!Ruc 'Ruc
'pR!Tipo_Cred  'tipo de Credito
'pR!Tipo_Prod  'tipo de producto
'lnMontoRiesgoUnico ' monto riesgo o prestamo
'****************************************************************
    'DERECHA
        'ETIQUETAS
            oDoc.WTextBox 80, 350, 12, 450, "Nº Crédito:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 88, 350, 12, 450, "Modalidad", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 96, 350, 12, 450, "Cod. Analista:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 104, 350, 12, 450, "Agencia:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 112, 350, 12, 450, "Línea de Crédito:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 120, 350, 12, 450, "Giro/Act Negocio:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 128, 350, 12, 450, "Direc. Principal:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 136, 350, 12, 450, "Direc. Negocio:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 144, 350, 12, 450, "Opinión Riesgo:", "F3", lnFontSizeBody, hLeft, , , , , , 2
        'DATOS
            oDoc.WTextBox 80, 400, 12, 450, pR!Nro_Credito, "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 88, 400, 12, 450, UCase(pR!Modalidad), "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 96, 400, 12, 450, UCase(pR!Analista), "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 104, 400, 12, 450, UCase(pR!Oficina), "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 112, 400, 12, 450, pR!Linea_Cred, "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 120, 400, 12, 450, pR!ActiGiro, "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 128, 400, 12, 450, pR!dire_domicilio, "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 136, 400, 12, 450, pR!dire_trabajo, "F3", lnFontSizeBody, hLeft, , , , , , 2
            If Not (prsOpRiesgo.EOF And prsOpRiesgo.BOF) Then
                oDoc.WTextBox 144, 400, 12, 450, prsOpRiesgo!cRiesgoValor, "F3", lnFontSizeBody, hLeft, , , , , , 2
            End If
    'FIN SECCION Nº 1
    'SECCION Nº 2
        oDoc.WTextBox 163, 55, 60, 490, "DATOS DEL CRÉDITO", "F4", lnFontSizeBody, hLeft
        oDoc.WTextBox 172, 55, 30, 510, "", "F3", lnFontSizeBody, hLeft, , , 1, vbBlack
    'IZQUIERDA
        'ETIQUETAS
            oDoc.WTextBox 172, 57, 12, 450, "Monto Propuesto:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 181, 57, 12, 450, "Plazo de Gracia:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 190, 57, 12, 450, "Plazo de Crédito:", "F3", lnFontSizeBody, hLeft, , , , , , 2
        'DATOS
            oDoc.WTextBox 172, 140, 12, 450, Format(pR!Ptmo_Propto, gcFormView), "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 181, 140, 12, 450, pR!PeriodoGracia, "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 190, 140, 12, 450, pR!Plazo, "F3", lnFontSizeBody, hLeft, , , , , , 2
    'CENTRO
        'ETIQUETAS
            oDoc.WTextBox 172, 280, 12, 450, "Tasa Propuesta(TEM):", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 181, 280, 12, 450, "Tasa Anterior (TEM):", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 190, 280, 12, 450, "Destino de Crédito:", "F3", lnFontSizeBody, hLeft, , , , , , 2
        'DATOS
            oDoc.WTextBox 172, 360, 12, 450, Format(pR!Tasa_Interes, gcFormDato), "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 181, 360, 12, 450, Format(pR!TEMAnterior, gcFormDato), "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 190, 360, 12, 450, pR!cDestino, "F3", lnFontSizeBody, hLeft, , , , , , 2
            'ALPA 20141030 *********************
            If CInt(pR!bExononeracionTasa) = 1 Then
                oDoc.WTextBox 210, 360, 12, 450, "EXONERADO DE TASA", "F3", lnFontSizeBody, hLeft, , , , , , 2
            End If
            'ALPA FIN **************************
    'DERECHA
        'ETIQUETAS
            oDoc.WTextBox 172, 400, 12, 450, "TEA Prop:", "F3", 7, hLeft, , , , , , 2
            oDoc.WTextBox 181, 400, 12, 450, "TEA Ant:", "F3", 7, hLeft, , , , , , 2
        'DATOS
            oDoc.WTextBox 172, 460, 12, 450, Round(((((pR!Tasa_Interes / 100) + 1) ^ (360 / 30)) - 1) * 100, 2), "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 181, 460, 12, 450, Round(((((pR!TEMAnterior / 100) + 1) ^ (360 / 30)) - 1) * 100, 2), "F3", lnFontSizeBody, hLeft, , , , , , 2
    'FIN SECCION Nº 2
    'SECCION Nº 3
        oDoc.WTextBox 206, 55 + 2, 60, 490, "GARANTIAS", "F4", lnFontSizeBody, hLeft
        'CABECERA
            oDoc.WTextBox 215, 57, 12, 36, "COD", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox 215, 93, 12, 60, "Tipo de Garantia", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox 215, 153, 12, 15, "RG", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox 215, 188 - 20, 12, 76, "Documento", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox 215, 264 - 20, 12, 76, "Dirección", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox 215, 340 - 20, 12, 20, "Mon.", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox 215, 367 - 27, 12, 40, "Val. Come.", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox 215, 407 - 27, 12, 40, "Val. Real.", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox 215, 447 - 27, 12, 40, "Val. Utilz.", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox 215, 487 - 27, 12, 40, "Val. Disp.", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox 215, 527 - 27, 12, 40, "Val. Grav.", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox 215, 567 - 27, 12, 25, "PROP.", "F1", nFTablaCabecera, hCenter, , , 1, , , 2

            Dim nCanGarant As Integer

            Dim nValComer As Double
            Dim nValReali As Double
            Dim nValUtiliz As Double
            Dim nValDisp As Double
            Dim nValGrava As Double


            nPosicion = 227

            If Not (prsCredGarant.BOF And prsCredGarant.EOF) Then
                If prsCredGarant.RecordCount > 0 Then
                'psRiesgo = "Riesgo 2" 'Comentado por JOEP 20170420
                
                    For i = 1 To prsCredGarant.RecordCount
                        Dim nAltoAdic As Integer
                        Dim nValorMayor As Integer
                        
                        'RECO20140804****************************
                        Dim nValCom As Double
                        Dim nValRea As Double
                        Dim nValUti As Double
                        Dim nValDis As Double
                        Dim nValGra As Double
                        
                        nValCom = IIf(prsCredGarant!cMoneda = "ME", prsCredGarant!nTasacion * pnTpoCambio, prsCredGarant!nTasacion)
                        nValRea = IIf(prsCredGarant!cMoneda = "ME", prsCredGarant!nRealizacion * pnTpoCambio, prsCredGarant!nRealizacion)
                        nValUti = IIf(prsCredGarant!cMoneda = "ME", prsCredGarant!nGravado * pnTpoCambio, prsCredGarant!nGravado)
                        nValDis = IIf(prsCredGarant!cMoneda = "ME", prsCredGarant!nDisponible * pnTpoCambio, prsCredGarant!nDisponible)
                        nValGra = IIf(prsCredGarant!cMoneda = "ME", prsCredGarant!nValorGravado * pnTpoCambio, prsCredGarant!nValorGravado)
                        'RECO FIN********************************
                        
                        nValorMayor = ValorMayor(Len(prsCredGarant!cTpoGarant), Len(prsCredGarant!cClasGarant), Len(prsCredGarant!cDocDesc), Len(prsCredGarant!cDireccion))

                        nAltoAdic = (nValorMayor / 13) * 6
                        
                        'INICIO Agrego JOEP 20170420
                            If prsCredGarant!nTratLegal = 1 Then
                                psRiesgo = "Riesgo 2"
                                nValorR2 = nValorR2 + 1
                            Else
                                psRiesgo = "Riesgo 1"
                                nValorR1 = nValorR1 + 1
                            End If
                        'FIN Agrego JOEP 20170420
                        
                        'Inicio Comentado por JOEP 20170420
                            'If Trim(prsCredGarant!cClasGarant) = "GARANTIAS NO PREFERIDAS" Then
                                'psRiesgo = "Riesgo 1"
                            'End If
                        'Fin Comentado por JOEP 20170420
                        
                        oDoc.WTextBox nPosicion + a, 57, 12 + nAltoAdic, 36, prsCredGarant!cNumGarant, "F1", nFTabla, hCenter, , , 1, , , 2
                        oDoc.WTextBox nPosicion + a, 93, 12 + nAltoAdic, 60, prsCredGarant!cTpoGarant, "F1", nFTabla, hjustify, , , 1, , , 2
                                               
                        oDoc.WTextBox nPosicion + a, 153, 12 + nAltoAdic, 15, IIf(prsCredGarant!nTratLegal = 1, 2, 1), "F1", nFTabla, hCenter, , , 1, , , 2 'Agrego JOEP 20170420
                                                
                        'oDoc.WTextBox nPosicion + A, 153, 12 + nAltoAdic, 15, IIf(Trim(prsCredGarant!cClasGarant) = "GARANTIAS NO PREFERIDAS", 1, 2), "F1", nFTabla, hCenter, , , 1, , , 2 'Comentado por JOEP 20170420
                        
                        oDoc.WTextBox nPosicion + a, 188 - 20, 12 + nAltoAdic, 76, prsCredGarant!cDocDesc & " - Nº " & prsCredGarant!cNroDoc, "F1", nFTabla, hjustify, , , 1, , , 2
                        'oDoc.WTextBox nPosicion + a, 264 - 20, 12 + nAltoAdic, 76, prsCredGarant!cDireccion, "F1", nFTabla, hjustify, , , 1, , , 2
                        oDoc.WTextBox nPosicion + a, 264 - 20, 12 + nAltoAdic, 76, prsCredGarant!cDireccion, "F1", nFTabla, hLeft, , , 1, , , 2 'FRHU 20150306 OBSERVACION
                        oDoc.WTextBox nPosicion + a, 340 - 20, 12 + nAltoAdic, 20, prsCredGarant!cMoneda, "F1", nFTabla, hCenter, , , 1, , , 2
                        'RECO20140804******************************************
                        oDoc.WTextBox nPosicion + a, 367 - 27, 12 + nAltoAdic, 40, Format(nValCom, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
                        oDoc.WTextBox nPosicion + a, 407 - 27, 12 + nAltoAdic, 40, Format(nValRea, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
                        oDoc.WTextBox nPosicion + a, 447 - 27, 12 + nAltoAdic, 40, Format(nValUti, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
                        oDoc.WTextBox nPosicion + a, 487 - 27, 12 + nAltoAdic, 40, Format(nValDis, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
                        oDoc.WTextBox nPosicion + a, 527 - 27, 12 + nAltoAdic, 40, Format(nValGra, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
                        'oDoc.WTextBox nPosicion + a, 367 - 27, 12 + nAltoAdic, 40, Format(prsCredGarant!nTasacion, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
                        'oDoc.WTextBox nPosicion + a, 407 - 27, 12 + nAltoAdic, 40, Format(prsCredGarant!nRealizacion, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
                        'oDoc.WTextBox nPosicion + a, 447 - 27, 12 + nAltoAdic, 40, Format(prsCredGarant!nGravado, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
                        'oDoc.WTextBox nPosicion + a, 487 - 27, 12 + nAltoAdic, 40, Format(prsCredGarant!ndisponible, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
                        'oDoc.WTextBox nPosicion + a, 527 - 27, 12 + nAltoAdic, 40, Format(prsCredGarant!nValorGravado, gcFormView), "F1", nFTabla, hRight, , , 1, , 2
                        'RECO FIN**********************************************
                        oDoc.WTextBox nPosicion + a, 567 - 27, 12 + nAltoAdic, 25, Mid(prsCredGarant!cRelGarant, 1, 3) & ".", "F1", nFTabla, hCenter, , , 1, , , 2
                        'RECO20140804*********************************
                        nValComer = nValComer + nValCom
                        nValReali = nValReali + nValRea
                        nValUtiliz = nValUtiliz + nValUti
                        nValDisp = nValDisp + nValDis
                        nValGrava = nValGrava + nValGra
                        'nValComer = nValComer + prsCredGarant!nTasacion
                        'nValReali = nValReali + prsCredGarant!nRealizacion
                        'nValUtiliz = nValUtiliz + prsCredGarant!nGravado
                        'nValDisp = nValDisp + prsCredGarant!ndisponible
                        'nValGrava = nValGrava + prsCredGarant!nValorGravado
                        'RECO FIN*************************************
                        prsCredGarant.MoveNext
                        a = a + 12 + nAltoAdic
                    Next
                End If
            End If
            'nPosicion = 207
            oDoc.WTextBox nPosicion + a, 57, 12, 283, "TOTALES", "F1", nFTabla, hCenter, , , 1, , , 2
            oDoc.WTextBox nPosicion + a, 340, 12, 40, Format(nValComer, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
            oDoc.WTextBox nPosicion + a, 380, 12, 40, Format(nValReali, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
            oDoc.WTextBox nPosicion + a, 420, 12, 40, Format(nValUtiliz, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
            oDoc.WTextBox nPosicion + a, 460, 12, 40, Format(nValDisp, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
            oDoc.WTextBox nPosicion + a, 500, 12, 40, Format(nValGrava, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
            oDoc.WTextBox nPosicion + a, 540, 12, 25, "", "F1", nFTabla, hLeft, , , 1, , , 2
            nPosicion = nPosicion + a + 20
            i = 0
            a = 0
    'FIN SECCION Nº 3
    'SECCION Nº 4
    'FIN SECCION Nº 4
    'SECCION Nº 5
        
        'RECO20140619************************
            Dim nPosAmp As Integer
            oDoc.WTextBox nPosicion + 2, 55, 60, 490, "COBERTURA DE GARANTIA", "F4", lnFontSizeBody, hLeft
            nPosAmp = nPosicion
            nPosicion = nPosicion + 12
            
            oDoc.WTextBox nPosicion, 55, 12, 95, "Cobertura Exp. Este Crédito", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox nPosicion, 150, 12, 95, "Cobertura Exp. Riesgo Único", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox nPosicion, 245, 12, 95, "Tipo de Riesgo", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            nPosicion = nPosicion + 12
            If lnMontoExpEstCred = 0 Then
                oDoc.WTextBox nPosicion + a, 55, 12, 95, Format(0, gcFormView), "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            Else
                'oDoc.WTextBox nPosicion + a, 55, 12, 95, Format(nValDisp / lnMontoExpEstCred, gcFormView), "F1", nFTablaCabecera, hCenter, , , 1, , , 2
                oDoc.WTextBox nPosicion + a, 55, 12, 95, Format(nValReali / lnMontoExpEstCred, gcFormView), "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            End If
            If lnMontoRiesgoUnico = 0 Then
                oDoc.WTextBox nPosicion + a, 150, 12, 95, Format(0, gcFormView), "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            Else
                'oDoc.WTextBox nPosicion + a, 150, 12, 95, Format(nValDisp / lnMontoRiesgoUnico, gcFormView), "F1", nFTablaCabecera, hCenter, , , 1, , , 2
                oDoc.WTextBox nPosicion + a, 150, 12, 95, Format(nValReali / lnMontoRiesgoUnico, gcFormView), "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            End If
            
            'INICION JOEP 20170420
            If (nValorR1 >= 1) Then
                nRunico = 1
                oDoc.WTextBox nPosicion + a, 245, 12, 95, "RIESGO 1", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            Else
                oDoc.WTextBox nPosicion + a, 245, 12, 95, UCase(psRiesgo), "F1", nFTablaCabecera, hCenter, , , 1, , , 2
                                
                'CTI3 (ferimoro)
                If psRiesgo = "Riesgo 1" Or psRiesgo = "RIESGO 1" Then
                    nRunico = 1
                Else
                  If psRiesgo = "Riesgo 2" Or psRiesgo = "RIESGO 2" Then
                    nRunico = 2
                  End If
                End If
                                
            End If
          'FIN JOEP 20170420
          
            'oDoc.WTextBox nPosicion + A, 245, 12, 95, UCase(psRiesgo), "F1", nFTablaCabecera, hCenter, , , 1, , , 2 'Comentado por JOEP 20170420
            
            nPosicion = nPosicion + 25
        'END RECO NUEVO********************
            
        If Not (prsCredAmp.BOF And prsCredAmp.EOF) Then
            oDoc.WTextBox nPosAmp + 2, 360, 60, 490, "AMPLIACIÓN DE CRÉDITO", "F4", lnFontSizeBody, hLeft
            
            
            nPosAmp = nPosAmp + 12
            oDoc.WTextBox nPosAmp, 360, 12, 77, "Crédito Nº", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox nPosAmp, 437, 12, 60, "Saldo Capital", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            nPosAmp = nPosAmp + 12
            For i = 1 To prsCredAmp.RecordCount
                oDoc.WTextBox nPosAmp + a, 360, 12, 77, prsCredAmp!cCtaCodAmp, "F1", nFTablaCabecera, hCenter, , , 1, , , 2
                oDoc.WTextBox nPosAmp + a, 437, 12, 60, Format(prsCredAmp!nMonto, gcFormView), "F1", nFTablaCabecera, hRight, , , 1, , , 2
                prsCredAmp.MoveNext
                a = a + 12
            Next
            nPosicion = nPosAmp + a + 10
        End If
    'FIN SECCION Nº 5
    'SECCION Nº 6
        Dim nCanIFs As Integer
        Dim nAdicionaFila As Integer
        If Not (pRRelaBcos.EOF And pRRelaBcos.BOF) Then
            nCanIFs = pRRelaBcos.RecordCount
        End If
        If nCanIFs > 5 Then
            nAdicionaFila = nAdicionaFila + 15
        End If
        If nCanIFs > 7 Then
            nAdicionaFila = nAdicionaFila + 15
        End If
        If nCanIFs > 9 Then
            nAdicionaFila = nAdicionaFila + 15
        End If
        nPosicion = nPosicion
        oDoc.WTextBox nPosicion + 6, 55, 60, 490, "RATIOS FINANCIEROS", "F4", lnFontSizeBody, hLeft
        oDoc.WTextBox nPosicion + 16, 55, 56 + nAdicionaFila, 240, "", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 16, 295, 56 + nAdicionaFila, 240, "", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 16, 295, 12, 139, "Institución", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 16, 434, 12, 28, "Moneda", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 16, 462, 12, 40, "Saldo", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 16, 502, 12, 33, "Relacion", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        nPosicion = nPosicion + 9
        oDoc.WTextBox nPosicion + 12, 57, 12, 63, "Liquidez", "F1", 5, hLeft, , , , , , 2
        'ENDEUDADAMIENTO CON OTRAS IFIS**************************************************
            Dim nIndx As Integer
            Dim nTmpPosic As Integer
            nTmpPosic = nPosicion + 9 '8
            For nIndx = 1 To pRRelaBcos.RecordCount
                'If Len(pRRelaBcos!Nombre) > 30 Then 'COMENTADO POR APRI 20170719 - MEJORA
                    oDoc.WTextBox nTmpPosic + 9, 295, 9, 139, pRRelaBcos!Nombre, "F1", 5, hLeft, , , , , , 2
                'Else 'COMENTADO POR APRI 20170719 - MEJORA
                    'oDoc.WTextBox nTmpPosic + 9, 295, 9, 139, pRRelaBcos!Nombre, "F1", nFTablaCabecera, hLeft, , , , , , 2 'COMENTADO POR APRI 20170719 - MEJORA
                'End If 'COMENTADO POR APRI 20170719 - MEJORA
                oDoc.WTextBox nTmpPosic + 9, 434, 9, 28, pRRelaBcos!Moneda, "F1", 5, hCenter, , , , , , 2
                oDoc.WTextBox nTmpPosic + 9, 462, 9, 40, Format(pRRelaBcos!Saldo, gcFormView), "F1", 5, hRight, , , , , , 2
                oDoc.WTextBox nTmpPosic + 9, 502, 9, 33, Mid(pRRelaBcos!Relacion, 1, 3) & ".", "F1", 5, hLeft, , , , , , 2
                'pRRelaBcos.MoveNext
                'APRI 20170719 - MEJORA
                If Len(pRRelaBcos!Nombre) > 42 Then
                nTmpPosic = nTmpPosic + 12
                Else
                 nTmpPosic = nTmpPosic + 8
                End If
                pRRelaBcos.MoveNext
                'END APRI
            Next
        'End If
        '********************************************************************************
'        oDoc.WTextBox nPosicion + 12, 120, 12, 63, Format(lnliquidez, gcFormView), "F1", nFTablaCabecera, hRight, , , 1, , , 2
'        oDoc.WTextBox nPosicion + 12, 188, 12, 63, "Patrimonio", "F1", 5, hLeft, , , , , , 2
'        oDoc.WTextBox nPosicion + 12, 221, 12, 63, Format(lnPatrimonio, gcFormView), "F1", nFTablaCabecera, hRight, , , 1, , , 2
'        nPosicion = nPosicion + 12
'        oDoc.WTextBox nPosicion + 12, 57, 12, 63, "Endeudamiento Patrimonial", "F1", 5, hLeft, , , , , , 2
'        oDoc.WTextBox nPosicion + 12, 120, 12, 63, Format(lnEndeudamiento * 100, "0.00") & "%", "F1", nFTablaCabecera, hRight, , , 1, , , 2
'        oDoc.WTextBox nPosicion + 12, 188, 12, 63, "Inventario", "F1", 5, hLeft, , , , , , 2
'        oDoc.WTextBox nPosicion + 12, 221, 12, 63, Format(lnInventario, gcFormView), "F1", nFTablaCabecera, hRight, , , 1, , , 2
'        nPosicion = nPosicion + 12
'        oDoc.WTextBox nPosicion + 12, 57, 12, 63, "Rentabilidad Patrimonial", "F1", 5, hLeft, , , , , , 2
'        oDoc.WTextBox nPosicion + 12, 120, 12, 63, Format(lnRentabPatrimonial * 100, "0.00") & "%", "F1", nFTablaCabecera, hRight, , , 1, , , 2
'        oDoc.WTextBox nPosicion + 12, 188, 12, 63, "Excedente", "F1", 5, hLeft, , , , , , 2
'        oDoc.WTextBox nPosicion + 12, 221, 12, 63, Format(lnExcedente, gcFormView), "F1", nFTablaCabecera, hRight, , , 1, , , 2
'        nPosicion = nPosicion + 12
'        oDoc.WTextBox nPosicion + 12, 57, 12, 63, "Capacid. De Pago", "F1", 5, hLeft, , , , , , 2
'        oDoc.WTextBox nPosicion + 12, 120, 12, 63, Format(lnCapacidadPago * 100, "0.00") & "%", "F1", nFTablaCabecera, hRight, , , 1, , , 2
        'DATOS DE RATIOS FORMATOS EVALUACION *************************************************************************************************
        oDoc.WTextBox nPosicion + 12, 120, 12, 63, IIf(pDRRatiosF!nLiquidezCte = 0, "N/A", Format(pDRRatiosF!nLiquidezCte, gcFormView)), "F1", nFTablaCabecera, hRight, , , 1, , , 2
        oDoc.WTextBox nPosicion + 12, 188, 12, 63, "Patrimonio", "F1", 5, hLeft, , , , , , 2
        oDoc.WTextBox nPosicion + 12, 221, 12, 63, IIf(pDRRatiosF!nPatrimonio = 0, "N/A", Format(pDRRatiosF!nPatrimonio, gcFormView)), "F1", nFTablaCabecera, hRight, , , 1, , , 2
        nPosicion = nPosicion + 12
        oDoc.WTextBox nPosicion + 12, 57, 12, 63, "Endeudamiento Patrimonial", "F1", 5, hLeft, , , , , , 2
        oDoc.WTextBox nPosicion + 12, 120, 12, 63, IIf(pDRRatiosF!nEndeuPat = 0, "N/A", Format(pDRRatiosF!nEndeuPat * 100, "0.00") & "%"), "F1", nFTablaCabecera, hRight, , , 1, , , 2
        oDoc.WTextBox nPosicion + 12, 188, 12, 63, "Ingreso N.", "F1", 5, hLeft, , , , , , 2
        oDoc.WTextBox nPosicion + 12, 221, 12, 63, IIf(pDRRatiosF!nIngreNeto = 0, "N/A", Format(pDRRatiosF!nIngreNeto, gcFormView)), "F1", nFTablaCabecera, hRight, , , 1, , , 2
        nPosicion = nPosicion + 12
        oDoc.WTextBox nPosicion + 12, 57, 12, 63, "Rentabilidad Patrimonial", "F1", 5, hLeft, , , , , , 2
        oDoc.WTextBox nPosicion + 12, 120, 12, 63, IIf(pDRRatiosF!nRentaPatri = 0, "N/A", Format(pDRRatiosF!nRentaPatri * 100, "0.00") & "%"), "F1", nFTablaCabecera, hRight, , , 1, , , 2
        oDoc.WTextBox nPosicion + 12, 188, 12, 63, "Excedente", "F1", 5, hLeft, , , , , , 2
        oDoc.WTextBox nPosicion + 12, 221, 12, 63, IIf(pDRRatiosF!nExceMensual = 0, "N/A", Format(pDRRatiosF!nExceMensual, gcFormView)), "F1", nFTablaCabecera, hRight, , , 1, , , 2
        nPosicion = nPosicion + 12
        oDoc.WTextBox nPosicion + 12, 57, 12, 63, "Capacid. De Pago", "F1", 5, hLeft, , , , , , 2
        oDoc.WTextBox nPosicion + 12, 120, 12, 63, IIf(pDRRatiosF!nCapPagNeta = 0, "N/A", Format(pDRRatiosF!nCapPagNeta * 100, "0.00") & "%"), "F1", nFTablaCabecera, hRight, , , 1, , , 2
        'FIN DATOS RATIOS FORMATOS EVALUACION *************************************************************************************************
        oDoc.WTextBox nPosicion + 12, 188, 12, 63, "Cuota", "F1", 5, hLeft, , , , , , 2
        oDoc.WTextBox nPosicion + 12, 221, 12, 63, Format(lnCuota, gcFormView), "F1", nFTablaCabecera, hRight, , , 1, , , 2
        nPosicion = nPosicion + 12 + 20 + nAdicionaFila 'aqui+ 20
    'FIN SECCION Nº 6
    'SECCION Nº 7
        oDoc.WTextBox nPosicion, 55, 60, 490, "CALIFICACIÓN Y RELACIÓN DE TITULARES / CÓNYUGE / AVALES", "F4", 7, hLeft
        oDoc.WTextBox nPosicion + 9, 55, 56, 480, "", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 9, 55, 12, 240, "Nombre", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 9, 295, 12, 115, "Relación", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 9, 410, 12, 25, "Normal", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 9, 435, 12, 25, "Poten.", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 9, 460, 12, 25, "Defic.", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 9, 485, 12, 25, "Dudos", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 9, 510, 12, 25, "Pérdida", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        'CALIFICACION Y RELACION DE TITULARES/CONYUGUE /AVALES**************************************************
        If Not (prsCalfSBSRela.EOF And prsCalfSBSRela.BOF) Then
            Dim nIndx2 As Integer
            Dim nTmpPosic2 As Integer
            nTmpPosic2 = nPosicion + 14
            For nIndx2 = 1 To prsCalfSBSRela.RecordCount
                oDoc.WTextBox nTmpPosic2 + 9, 55, 11, 240, prsCalfSBSRela!cPersNombre, "F1", nFTablaCabecera, hLeft, , , , , , 1
                oDoc.WTextBox nTmpPosic2 + 9, 295, 11, 115, prsCalfSBSRela!cConsDescripcion, "F1", nFTablaCabecera, hLeft, , , , , , 1
                '*** FRHU 20160823
                'oDoc.WTextBox nTmpPosic2 + 9, 410, 11, 25, prsCalfSBSRela!Normal & "%", "F1", nFTablaCabecera, hCenter, , , , , , 1
                'oDoc.WTextBox nTmpPosic2 + 9, 435, 11, 25, prsCalfSBSRela!Potencial & "%", "F1", nFTablaCabecera, hCenter, , , , , , 1
                'oDoc.WTextBox nTmpPosic2 + 9, 460, 11, 25, prsCalfSBSRela!DEFICIENTE & "%", "F1", nFTablaCabecera, hCenter, , , , , , 1
                'oDoc.WTextBox nTmpPosic2 + 9, 485, 11, 25, prsCalfSBSRela!DUDOSO & "%", "F1", nFTablaCabecera, hCenter, , , , , , 1
                'oDoc.WTextBox nTmpPosic2 + 9, 510, 11, 25, prsCalfSBSRela!PERDIDA & "%", "F1", nFTablaCabecera, hCenter, , , , , , 1
                oDoc.WTextBox nTmpPosic2 + 9, 410, 11, 28, prsCalfSBSRela!Normal & "%", "F1", nFTablaCabecera, hCenter, , , , , , 1
                oDoc.WTextBox nTmpPosic2 + 9, 435, 11, 28, prsCalfSBSRela!Potencial & "%", "F1", nFTablaCabecera, hCenter, , , , , , 1
                oDoc.WTextBox nTmpPosic2 + 9, 460, 11, 28, prsCalfSBSRela!DEFICIENTE & "%", "F1", nFTablaCabecera, hCenter, , , , , , 1
                oDoc.WTextBox nTmpPosic2 + 9, 485, 11, 28, prsCalfSBSRela!DUDOSO & "%", "F1", nFTablaCabecera, hCenter, , , , , , 1
                oDoc.WTextBox nTmpPosic2 + 9, 510, 11, 28, prsCalfSBSRela!PERDIDA & "%", "F1", nFTablaCabecera, hCenter, , , , , , 1
                '*** FIN FRHU 20160823
                prsCalfSBSRela.MoveNext
                nTmpPosic2 = nTmpPosic2 + 8
            Next
        End If
    'FIN SECCION Nº 7
    'APRI20170719 TI-ERS025-2017
    If prsPersVinc.RecordCount > 0 Then
        nPosicion = nPosicion + 70
        oDoc.WTextBox nPosicion, 55, 60, 490, "CALIFICACIÓN DE VINCULADOS AL COLABORADOR", "F4", 7, hLeft
        oDoc.WTextBox nPosicion + 9, 55, 56, 480, "", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 9, 55, 12, 240, "Nombre", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 9, 295, 12, 115, "Relación", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 9, 410, 12, 25, "Normal", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 9, 435, 12, 25, "Poten.", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 9, 460, 12, 25, "Defic.", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 9, 485, 12, 25, "Dudos", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 9, 510, 12, 25, "Pérdida", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        If Not (prsPersVinc.EOF And prsPersVinc.BOF) Then
            Dim nIndxP As Integer
            Dim nTmpPosicP As Integer
            nTmpPosicP = nPosicion + 14
            For nIndxP = 1 To prsPersVinc.RecordCount
                oDoc.WTextBox nTmpPosicP + 9, 55, 11, 240, prsPersVinc!cPersNombre, "F1", nFTablaCabecera, hLeft, , , , , , 1
                oDoc.WTextBox nTmpPosicP + 9, 295, 11, 115, prsPersVinc!cRelacion, "F1", nFTablaCabecera, hLeft, , , , , , 1
                oDoc.WTextBox nTmpPosicP + 9, 410, 11, 28, prsPersVinc!Normal & "%", "F1", nFTablaCabecera, hCenter, , , , , , 1
                oDoc.WTextBox nTmpPosicP + 9, 435, 11, 28, prsPersVinc!Potencial & "%", "F1", nFTablaCabecera, hCenter, , , , , , 1
                oDoc.WTextBox nTmpPosicP + 9, 460, 11, 28, prsPersVinc!DEFICIENTE & "%", "F1", nFTablaCabecera, hCenter, , , , , , 1
                oDoc.WTextBox nTmpPosicP + 9, 485, 11, 28, prsPersVinc!DUDOSO & "%", "F1", nFTablaCabecera, hCenter, , , , , , 1
                oDoc.WTextBox nTmpPosicP + 9, 510, 11, 28, prsPersVinc!PERDIDA & "%", "F1", nFTablaCabecera, hCenter, , , , , , 1
                prsPersVinc.MoveNext
                nTmpPosicP = nTmpPosicP + 8
            Next
        End If
    End If
    'END APRI
    'WIOR 20160621 *** SOBREENDEUDAMIENTO DE CLIENTES
    nPosicion = nPosicion + 70
    If Not (prsSobreEnd.EOF And prsSobreEnd.BOF) Then
        oDoc.WTextBox nPosicion, 55, 56, 330, "SOBREENDEUDAMIENTO DE CLIENTES", "F4", 7, hLeft
        nPosicion = nPosicion + 12
        oDoc.WTextBox nPosicion, 55, 12, 80, "Deuda Potencial", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion, 135, 12, 80, Format(CCur(prsSobreEnd!nDeudaTotal), "###," & String(15, "#") & "#0.00"), "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        If Not (prsSobreEndCodigos.EOF And prsSobreEndCodigos.BOF) Then
            nPosicion = nPosicion + 18

            oDoc.WTextBox nPosicion, 55, 12, 50, "Códigos", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox nPosicion, 105, 12, 110, "Resultados", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox nPosicion, 215, 12, 160, "Detalle", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox nPosicion, 375, 12, 160, "Plan de Mitigación del Riesgo", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            nPosicion = nPosicion + 12

            For i = 1 To prsSobreEndCodigos.RecordCount
                '*** FRHU 20160823
                'oDoc.WTextBox nPosicion, 55, 15, 50, Trim(prsSobreEndCodigos!cCodigo), "F1", nFTablaCabecera, hCenter, vMiddle, , 1, , , 2
                'oDoc.WTextBox nPosicion, 105, 15, 110, Trim(prsSobreEndCodigos!cResultado), "F1", nFTablaCabecera, hLeft, vMiddle, , 1, , , 2
                'oDoc.WTextBox nPosicion, 215, 15, 160, Trim(prsSobreEndCodigos!cDetalle), "F1", 6, hLeft, vMiddle, , 1, , , 2
                'oDoc.WTextBox nPosicion, 375, 15, 160, Trim(prsSobreEndCodigos!cPlanMitigacion), "F1", 6, hLeft, vMiddle, , 1, , , 2
                'prsSobreEndCodigos.MoveNext
                'nPosicion = nPosicion + 15
                oDoc.WTextBox nPosicion, 55, 22, 50, Trim(prsSobreEndCodigos!cCodigo), "F1", nFTablaCabecera, hCenter, vMiddle, , 1, , , 2
                oDoc.WTextBox nPosicion, 105, 22, 110, Trim(prsSobreEndCodigos!cResultado), "F1", nFTablaCabecera, hLeft, vMiddle, , 1, , , 2
                oDoc.WTextBox nPosicion, 215, 22, 160, Trim(prsSobreEndCodigos!cDetalle), "F1", 6, hLeft, , , 1, , , 2
                oDoc.WTextBox nPosicion, 375, 22, 160, Trim(prsSobreEndCodigos!cPlanmitigacion), "F1", 6, hLeft, vMiddle, , 1, , , 2
                prsSobreEndCodigos.MoveNext
                nPosicion = nPosicion + 22
                '*** FIN FRHU 20160823
            Next i
        End If
        'nPosicion = nPosicion + 12
        nPosicion = nPosicion + 4 'FRHU 20160823
    End If
    'WIOR FIN ********
    
    'RECO NUEVO
    'nPosicion = nPosicion + 70
    If Not (prsComentAnalis.BOF And prsComentAnalis.EOF) Then
        oDoc.WTextBox nPosicion, 55, 56, 330, "COMENTARIO ANALISTA", "F1", nFTablaCabecera, hLeft
        nPosicion = nPosicion + 9
        oDoc.WTextBox nPosicion, 55, 40, 480, prsComentAnalis!cComentAnalista, "F1", nFTablaCabecera, hjustify, , , 1, , , 2
        nPosicion = nPosicion + 48
    End If
    
    'RECO FIN NUEVO
    'SECCION Nº 8
        'oDoc.WTextBox nPosicion, 55, 56, 280, "", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion, 55, 100, 210, "", "F1", nFTablaCabecera, hCenter, , , 1, , , 2 'FRHU 20160811 Anexo002
        'oDoc.WTextBox nPosicion, 55, 56, 165, "", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        'oDoc.WTextBox nPosicion, 55, 12, 165, "EXONERACIONES", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        'oDoc.WTextBox nPosicion, 220, 12, 165, "AUTORIZACIONES", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        'oDoc.WTextBox nPosicion, 55, 12, 280, "AUTORIZACIONES", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion, 55, 12, 210, "Autorizaciones", "F1", nFTablaCabecera, hCenter, , , 1, , , 2 'FRHU 20160811 Anexo002
        'oDoc.WTextBox nPosicion, 335, 12, 200, "NIVELES DE APROBACION POR EXPOSICION", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion, 265, 12, 270, "Nivel de Aprobación", "F1", nFTablaCabecera, hCenter, , , 1, , , 2 'FRHU 20160811 Anexo002 ERS002-2016
        
        'oDoc.WTextBox nPosicion, 335, 56, 200, "", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion, 265, 100, 270, "", "F1", nFTablaCabecera, hCenter, , , 1, , , 2 'FRHU 20160811 Anexo002
        
        nPosicion = nPosicion + 12 'FRHU 20160811 Anexo002 ERS002-2016 NEW
        oDoc.WTextBox nPosicion, 265, 12, 135, "Por Autorización", "F1", nFTablaCabecera, hCenter, , , 1, , , 2 'FRHU 20160811 Anexo002 ERS002-2016 NEW
        oDoc.WTextBox nPosicion, 400, 12, 135, "Por Exposición", "F1", nFTablaCabecera, hCenter, , , 1, , , 2 'FRHU 20160811 Anexo002 ERS002-2016 NEW
        oDoc.WTextBox nPosicion, 400, 88, 135, "", "F1", nFTablaCabecera, hCenter, , , 1, , , 2 'FRHU 20160811 Anexo002
        
        If psHabilitaNiveles = 0 Then
            If Not (prsExoAutCred.EOF And prsExoAutCred.BOF) Then
                Dim nIndx3 As Integer
                Dim nTmpPosic3 As Integer
                Dim nTmpPosicExo As Integer
                Dim nTmpPosicAut As Integer
                nTmpPosic3 = nPosicion + 14
                nTmpPosicExo = nPosicion + 2
                nTmpPosicAut = nPosicion + 2
                'For nIndx3 = 1 To 3
                For nIndx3 = 1 To prsExoAutCred.RecordCount
                 Dim texto As String
                    If prsExoAutCred!nTipoExoneraCod = 1 Then
                        Set lrsRespNvlExo = oDCredExoAut.RecuperaRespExoAut(prsExoAutCred!cExoneraCod, IIf(lnMontoRiesgoUnico = 0, lnMontoExpEstCred, lnMontoRiesgoUnico))
                        oDoc.WTextBox nTmpPosicExo + 9, 55, 9, 150, prsExoAutCred!cExoneraDesc, "F1", nFTablaCabecera, hLeft, , , , , , 2
                        
                        If Not (lrsRespNvlExo.BOF And lrsRespNvlExo.EOF) Then
                            oDoc.WTextBox nTmpPosicExo + 9, 150, 9, 150, lrsRespNvlExo!cNivAprDesc, "F1", nFTablaCabecera, hLeft, , , , , , 2
                        End If
                        nTmpPosicExo = nTmpPosicExo + 5
                    Else
                        Set lrsRespNvlExo = oDCredExoAut.RecuperaRespExoAut(prsExoAutCred!cExoneraCod, IIf(lnMontoRiesgoUnico = 0, lnMontoExpEstCred, lnMontoRiesgoUnico))
                        oDoc.WTextBox nTmpPosicAut + 9, 225, 9, 240, prsExoAutCred!cExoneraDesc, "F1", nFTablaCabecera, hLeft, , , , , , 2
                        If Not (lrsRespNvlExo.BOF And lrsRespNvlExo.EOF) Then
                            oDoc.WTextBox nTmpPosicAut + 9, 300, 9, 150, lrsRespNvlExo!cNivAprDesc, "F1", nFTablaCabecera, hLeft, , , , , , 2
                        End If
                        nTmpPosicAut = nTmpPosicAut + 5
                    End If
                    prsExoAutCred.MoveNext
                    Set lrsRespNvlExo = Nothing
                Next
            End If
'***********************FERIMORO********************
'EL CODIGO DE ABAJO ES ORIGINAL

'            If Not (prsCredResulNivApr.EOF And prsCredResulNivApr.BOF) Then
'                Dim nIndx4 As Integer
'                Dim nTmpPosic4 As Integer
'
'                nTmpPosic4 = nPosicion + 14
'                For nIndx = 1 To prsCredResulNivApr.RecordCount
'                    oDoc.WTextBox nTmpPosic4 + 12, 390, 12, 305, prsCredResulNivApr!cNivAprDesc, "F1", nFTablaCabecera, hLeft
'                    nTmpPosic4 = nTmpPosic4 + 5
'                    prsCredResulNivApr.MoveNext
'                Next
'            End If

            '***********************FERIMORO********************
            '***********VISUALIZAR EL NIVEL DE APROBACIONN******
            'nTmpPosicNiv4 + 12,390  -555
            'pRNivApr
            '***************************************************

If pR!Modalidad = "REFINANCIADO" Then
 Dim sNivelinicial As String
 Dim sPdirecto
 Dim nTmpPosicNiv4 As Integer
     
     Dim oRecCtaRef As COMDCredito.DCOMCredito
     Dim cCtaRefe As ADODB.Recordset
 
    Set oRecCtaRef = New COMDCredito.DCOMCredito
    Set cCtaRefe = oRecCtaRef.RecuperaCctaReferencia(pR!Nro_Credito)
    Set oRecCtaRef = Nothing

    sNivelinicial = NivelAprobacionInicial(pR!Nro_Credito, cCtaRefe!cCtaCodRef)


    sPdirecto = pR!Tipo_Prod

    Set oDpExposicion = New COMDCredito.DCOMCredito
    Set expo = oDpExposicion.RecuperaAutoporExpo(nRunico, pR!cDestino, pR!Modalidad, lnMontoRiesgoUnico, pR!Oficina, sPdirecto, nClePref, sNivelinicial)
    Set oDpExposicion = Nothing
    
            nTmpPosicNiv4 = nPosicion + 1
            oDoc.WTextBox nTmpPosicNiv4 + 12, 402, 12, 305, IIf(IsNull(expo!autExp) = True, "", expo!autExp), "F1", nFTablaCabecera, hLeft
    
Else
                     '
    
    If pR!Tipo_Prod = "Agropecuario Directo" Then
    sPdirecto = "AGROPECUARIOS DIRECTO"
    Else
    sPdirecto = pR!Tipo_Prod
    End If
                                    
    Set oDpExposicion = New COMDCredito.DCOMCredito
    Set expo = oDpExposicion.RecuperaAutoporExpo(nRunico, pR!cDestino, pR!Modalidad, lnMontoRiesgoUnico, pR!Oficina, sPdirecto, nClePref)
    Set oDpExposicion = Nothing


            
            nTmpPosicNiv4 = nPosicion + 1
            oDoc.WTextBox nTmpPosicNiv4 + 12, 402, 12, 305, IIf(IsNull(expo!autExp) = True, "", expo!autExp), "F1", nFTablaCabecera, hLeft
End If
            
        End If
        'FRHU 20160811 Anexo002 ERS002-2016
                  '**********  niveles _ ferimoro
        If Not (prsAutorizaciones.EOF And prsAutorizaciones.BOF) Then
            Dim nTmpPosicExo5 As Integer
            nTmpPosicExo5 = nPosicion + 2
            For nIndx = 1 To prsAutorizaciones.RecordCount
                If Len(prsAutorizaciones!cExoneraDesc) <= 69 Then
                    oDoc.WTextBox nTmpPosicExo5 + 9, 55, 9, 280, prsAutorizaciones!cExoneraDesc, "F1", nFTablaCabecera, hLeft, , , , , , 2
                    oDoc.WTextBox nTmpPosicExo5 + 11, 267, 12, 305, prsAutorizaciones!cNivAprDesc, "F1", nFTablaCabecera, hLeft
                    nTmpPosicExo5 = nTmpPosicExo5 + 9
                Else
                    oDoc.WTextBox nTmpPosicExo5 + 9, 55, 9, 280, prsAutorizaciones!cExoneraDesc, "F1", nFTablaCabecera, hLeft, , , , , , 2
                    oDoc.WTextBox nTmpPosicExo5 + 11, 267, 12, 305, prsAutorizaciones!cNivAprDesc, "F1", nFTablaCabecera, hLeft
                    nTmpPosicExo5 = nTmpPosicExo5 + 18
                End If
                prsAutorizaciones.MoveNext
            Next nIndx
        End If
        'FIN FRHU 20160811
    'FIN SECCION Nº 8
    
    'SECCION Nº9
    'nPosicion = nPosicion + 50
    nPosicion = nPosicion + 80 'FRHU 20160811 Anexo002 ERS002-2016
    oDoc.WTextBox nPosicion + 15, 150, 56, 150, "RESOLUCION DE COMITÉ, EN CONCLUSION: ", "F1", nFTablaCabecera, Left
    nPosicion = nPosicion + 12
    oDoc.WTextBox nPosicion + 15, 150, 56, 70, "MONTO", "F1", nFTablaCabecera, Left
    oDoc.WTextBox nPosicion + 15, 250, 56, 70, "CUOTAS", "F1", nFTablaCabecera, Left
    oDoc.WTextBox nPosicion + 15, 350, 56, 70, "TI", "F1", nFTablaCabecera, Left
    oDoc.WTextBox nPosicion + 15, 450, 56, 70, "VCTO", "F1", nFTablaCabecera, Left
    
    nPosicion = nPosicion + 20
    oDoc.WTextBox nPosicion + 15, 70, 56, 150, "APROBADO POR: ", "F1", nFTablaCabecera, Left
    oDoc.WTextBox nPosicion + 15, 150, 56, 70, "...................", "F1", nFTablaCabecera, Left
    oDoc.WTextBox nPosicion + 15, 250, 56, 70, "...................", "F1", nFTablaCabecera, Left
    oDoc.WTextBox nPosicion + 15, 350, 56, 70, "...................", "F1", nFTablaCabecera, Left
    oDoc.WTextBox nPosicion + 15, 450, 56, 70, "...................", "F1", nFTablaCabecera, Left
    'FIN SECCION N9
    oDoc.PDFClose
    oDoc.Show
    Exit Sub
ErrorImprimirPDF:
    MsgBox Err.Description, vbInformation, "Aviso"
End Sub

Private Sub ImprimeHojaAprobacionCredLeasing(ByRef pR As ADODB.Recordset, ByRef pRB As ADODB.Recordset, ByRef pRCliRela As ADODB.Recordset, ByRef pRRelaBcos As ADODB.Recordset, _
                                      ByRef pRDatFinan As ADODB.Recordset, ByRef pRResGarTitAva As ADODB.Recordset, ByRef pRGarantCred As ADODB.Recordset, _
                                      ByVal psCodCta, ByVal prsCredEval As ADODB.Recordset, ByVal prsCredAmp As ADODB.Recordset, ByVal prsCalfSBSRela As ADODB.Recordset, _
                                      ByVal prsExoAutCred As ADODB.Recordset, ByVal prsRiesgoUnico As ADODB.Recordset, ByVal prsCredGarant As ADODB.Recordset, _
                                      ByVal prsCredResulNivApr As ADODB.Recordset, ByVal pRNivApr As ADODB.Recordset, ByVal psRiesgo As String, ByVal psHabilitaNiveles As String, _
                                      ByVal prsOpRiesgo As ADODB.Recordset, ByVal prsComentAnalis As ADODB.Recordset, _
                                      ByVal pnTpoCambio As Double, ByVal prsSobreEnd As ADODB.Recordset, ByVal prsSobreEndCodigos As ADODB.Recordset, _
                                      ByVal pDRRatiosF As ADODB.Recordset, ByVal prsAutorizaciones As ADODB.Recordset, ByVal prsPersVinc As ADODB.Recordset) 'RECO20140804 Se añadió el parametro pnTpoCambio
                                      'WIOR 20160621 AGREGO ByVal prsSobreEnd As ADODB.Recordset, ByVal prsSobreEndCodigos As ADODB.Recordset
                                      'RECO20160730 SE AGREGO RATIOS DE FORMATO EVALUACION
                                      'FRHU 20160811 Anexo002 ERS002-2016
                                      'APRI 20170719 AGREGO prsPersVinc TI-ERS-025 2017
    Dim oFun As New COMFunciones.FCOMImpresion
    Dim lnliquidez As Double, lnCapacidadPago As Double, lnExcedente As Double
    Dim lnPatriEmpre As Double, lnPatrimonio As Double, lnIngresoNeto As Double
    Dim lnRentabPatrimonial As Double, lnEndeudamiento As Double
    Dim lnInventario As Double, lnCuota As Double
    Dim lnMontoRiesgoUnico As Double, lnRemNeta As Double, lnEgresos As Double
    Dim lnMontoExpEstCred As Double
    'ALPA20160714****************************
    Dim lnCapacidadPago2 As Double
    '****************************************
    'Dim oCliPre As New COMNCredito.NCOMCredito 'COMENTADO POR ARLO 20170722
    Dim oDCredExoAut As New COMDCredito.DCOMNivelAprobacion
    Dim lrsRespNvlExo As ADODB.Recordset
    Dim oCrDoc As New COMDCredito.DCOMCredDoc
    Set oDCredExoAut = New COMDCredito.DCOMNivelAprobacion
    Dim bValidarCliPre As Boolean
    
    Dim lrDatosRatiosFI As ADODB.Recordset
    
    On Error GoTo ErrorImprimirPDF
    Dim oDoc  As cPDF
    Dim nTipo As Integer
    Set oDoc = New cPDF
    
    
    Dim i As Integer
    Dim a As Integer
    Dim nPosicion As Integer
    
    Dim dFecFteIng As Date
    
    If oCrDoc.ObtenerFechaProNumFuente(pR!cNumeroFuente, gdFecSis).RecordCount > 0 Then
        dFecFteIng = Format(oCrDoc.ObtenerFechaProNumFuente(pR!cNumeroFuente, gdFecSis)!valor, "dd/mm/yyyy")
    End If
    
    Set lrDatosRatiosFI = New ADODB.Recordset
    Set lrDatosRatiosFI = oCrDoc.ObtieneRatiosFI(pR!cNumeroFuente, psCodCta)
    
    lnliquidez = 0
    lnEndeudamiento = 0
    lnRentabPatrimonial = 0
    lnCapacidadPago = 0
    lnPatrimonio = 0
    lnInventario = 0
    lnExcedente = 0
    lnCuota = 0
    lnRemNeta = 0
    lnEgresos = 0
    
    nValorR1 = 0 'JOEP20170627
    nValorR2 = 0 'JOEP20170627
    
    If Not (lrDatosRatiosFI.EOF And lrDatosRatiosFI.BOF) Then
        Dim nIndexRatios As Integer
        For nIndexRatios = 0 To lrDatosRatiosFI.RecordCount - 1
            If lrDatosRatiosFI!nValor = 1 Then
                lnInventario = lrDatosRatiosFI!nMonto
            ElseIf lrDatosRatiosFI!nValor = 2 Then
                lnExcedente = lrDatosRatiosFI!nMonto
            ElseIf lrDatosRatiosFI!nValor = 3 Then
                lnCuota = lrDatosRatiosFI!nMonto
            ElseIf lrDatosRatiosFI!nValor = 4 Then
                lnPatrimonio = lrDatosRatiosFI!nMonto
            ElseIf lrDatosRatiosFI!nValor = 5 Then
                lnEndeudamiento = lrDatosRatiosFI!nMonto
            ElseIf lrDatosRatiosFI!nValor = 6 Then
                lnliquidez = lrDatosRatiosFI!nMonto
            'ALPA20160714*****************************************************************************
            ElseIf lrDatosRatiosFI!nValor = 11 Then
                lnCapacidadPago2 = IIf(IsNull(lrDatosRatiosFI!nMonto), 0, lrDatosRatiosFI!nMonto)
            '*****************************************************************************************
            End If
            lrDatosRatiosFI.MoveNext
        Next
        'ALPA20160714******************************************************************
        'lnCapacidadPago = lnCuota / IIf(lnExcedente = 0, 1, lnExcedente)
        If lnCapacidadPago2 = 0 Then
            lnCapacidadPago = lnCuota / IIf(lnExcedente = 0, 1, lnExcedente)
        Else
            lnCapacidadPago = lnCapacidadPago2
        End If
        '******************************************************************************
        lnRentabPatrimonial = IIf(lnExcedente = 0, 1, lnExcedente) / lnPatrimonio
        lnEndeudamiento = lnEndeudamiento / lnPatrimonio
    End If
    
    oDoc.Author = gsCodUser
    oDoc.Creator = "SICMACT - Negocio"
    oDoc.Producer = "Caja Municipal de Ahorros y Crédito de Maynas S.A."
    oDoc.Subject = "Hoja de Aprobación de Créditos Nº " & "RECO"
    oDoc.Title = "Hoja de Aprobación de Créditos Nº " & "RECO"
    If Not oDoc.PDFCreate(App.Path & "\Spooler\" & IIf(nTipo = 1, "Previo", "") & psCodCta & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
        Exit Sub
    End If

    'FUENTES
    Dim nFTabla As Integer
    Dim nFTablaCabecera As Integer
    Dim lnFontSizeBody As Integer
    
    oDoc.Fonts.Add "F1", "Times New Roman", TrueType, Normal, WinAnsiEncoding 'FRHU20131126
    oDoc.Fonts.Add "F2", "Times New Roman", TrueType, Bold, WinAnsiEncoding 'FRHU20131126
    oDoc.Fonts.Add "F3", "Arial Narrow", TrueType, Normal, WinAnsiEncoding 'RECO20140227
    oDoc.Fonts.Add "F4", "Arial Narrow", TrueType, Bold, WinAnsiEncoding 'RECO20140227
    
    nFTablaCabecera = 7
    nFTabla = 7
    lnFontSizeBody = 7
    'FIN FUENTES

    oDoc.NewPage A4_Vertical
    lnMontoRiesgoUnico = 0#
    If Not (prsRiesgoUnico.EOF And prsRiesgoUnico.BOF) Then
        lnMontoRiesgoUnico = prsRiesgoUnico!nMonto
    End If


            'Set oCliPre = New COMNCredito.NCOMCredito 'COMENTADO POR ARLO 20170722
            'bValidarCliPre = oCliPre.ValidarClientePreferencial(pR!cPersCod) 'COMENTADO POR ARLO 20170722
            bValidarCliPre = False 'ARLO 20170722
    
    'oDoc.WTextBox 40, 70, 10, 450, "HOJA DE APROBACIÓN DE LEASING:", "F3", 11, hCenter, , vbBlack
    'RECO20141011 ***********************************
    If bValidarCliPre = True Then
        oDoc.WTextBox 40, 70, 10, 450, "HOJA DE APROBACIÓN - CLIENTE PREFERENCIAL", "F3", 11, hCenter, , vbBlack
    Else
        oDoc.WTextBox 40, 70, 10, 450, "HOJA DE APROBACIÓN", "F3", 11, hCenter, , vbBlack
    End If
    'RECO END ***************************************
    oDoc.WTextBox 60, 55, 10, 200, "APROBACION DE ARRENDAMIENTO FIANCIERO", "F4", 9, hLeft
    oDoc.WTextBox 60, 354, 10, 200, "FECHA APROBACION:....................", "F4", 9, hRight

    'SECCION Nº 1
        'oDoc.WTextBox 80, 55, 60, 490, "", "F3", 7, hLeft, , , 1, vbBlack
        oDoc.WTextBox 80, 55, 80, 490, "", "F3", 7, hLeft, , , 1, vbBlack 'FRHU 20160909
    'IZQUIERDA
        'ETIQUETAS
            oDoc.WTextBox 80, 57, 12, 450, "Cliente:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 88, 57, 12, 450, "DNI/RUC:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 96, 57, 12, 450, "Tipo de Credito:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 104, 57, 12, 450, "Producto de Crédito:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 112, 57, 12, 450, "Exposición Anterior Máxima:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 120, 57, 12, 450, "Exposición con este Crédito:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 128, 57, 12, 450, "Exposición Riesgo Único:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 136, 57, 12, 450, "Fecha de Solicitud:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 144, 57, 12, 450, "Campaña:", "F3", lnFontSizeBody, hLeft, , , , , , 2
        'DATOS
            oDoc.WTextBox 80, 140, 12, 450, UCase(pR!Prestatario), "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 88, 140, 12, 450, pR!DniRuc & "/" & pR!Ruc, "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 96, 140, 12, 450, pR!Tipo_Cred, "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 104, 140, 12, 450, pR!Tipo_Prod, "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 112, 140, 12, 450, Format(pR!ExpoAntMax, gcFormView), "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 120, 140, 12, 450, Format(pR!ExpoConCred, gcFormView), "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 128, 140, 12, 450, Format(lnMontoRiesgoUnico, gcFormView), "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 136, 140, 12, 450, pR!Fec_Soli, "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 144, 145, 12, 450, pR!cCampania, "F3", lnFontSizeBody, hLeft, , , , , , 2
            lnMontoExpEstCred = pR!ExpoConCred

    'DERECHA
        'ETIQUETAS
            oDoc.WTextBox 80, 350, 12, 450, "Nº Crédito:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 88, 350, 12, 450, "Modalidad", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 96, 350, 12, 450, "Cod. Analista:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 104, 350, 12, 450, "Agencia:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 112, 350, 12, 450, "Línea de Crédito:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 120, 350, 12, 450, "Giro/Act Negocio:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 128, 350, 12, 450, "Direc. Principal:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 136, 350, 12, 450, "Direc. Negocio:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 144, 350, 12, 450, "Opinión Riesgo:", "F3", lnFontSizeBody, hLeft, , , , , , 2
        'DATOS
            oDoc.WTextBox 80, 400, 12, 450, pR!Nro_Credito, "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 88, 400, 12, 450, UCase(pR!Modalidad), "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 96, 400, 12, 450, UCase(pR!Analista), "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 104, 400, 12, 450, UCase(pR!Oficina), "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 112, 400, 12, 450, pR!Linea_Cred, "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 120, 400, 12, 450, pR!ActiGiro, "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 128, 400, 12, 450, pR!dire_domicilio, "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 136, 400, 12, 450, pR!dire_trabajo, "F3", lnFontSizeBody, hLeft, , , , , , 2
            If Not (prsOpRiesgo.EOF And prsOpRiesgo.BOF) Then
                oDoc.WTextBox 144, 400, 12, 450, prsOpRiesgo!cRiesgoValor, "F3", lnFontSizeBody, hLeft, , , , , , 2
            End If
    'FIN SECCION Nº 1
    'SECCION Nº 2
        oDoc.WTextBox 163, 55, 60, 490, "DATOS DEL CRÉDITO", "F4", lnFontSizeBody, hLeft
        oDoc.WTextBox 172, 55, 30, 510, "", "F3", lnFontSizeBody, hLeft, , , 1, vbBlack
    'IZQUIERDA
        'ETIQUETAS
            oDoc.WTextBox 172, 57, 12, 450, "Monto Propuesto:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 181, 57, 12, 450, "Plazo de Gracia:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 190, 57, 12, 450, "Plazo de Crédito:", "F3", lnFontSizeBody, hLeft, , , , , , 2
        'DATOS
            oDoc.WTextBox 172, 140, 12, 450, Format(pR!Ptmo_Propto, gcFormView), "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 181, 140, 12, 450, pR!PeriodoGracia, "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 190, 140, 12, 450, pR!Plazo, "F3", lnFontSizeBody, hLeft, , , , , , 2
        '*** FRHU 20160909
            Dim oLeasing As COMNCredito.NCOMLeasing
            Set oLeasing = New COMNCredito.NCOMLeasing
            oDoc.WTextBox 172, 180, 12, 450, "Monto Financiado:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 172, 230, 12, 450, Format(oLeasing.ObtenerMontoFinanciado(pR!Nro_Credito), gcFormView), "F3", lnFontSizeBody, hLeft, , , , , , 2
        '*** FIN FRHU 20160909
    'CENTRO
        'ETIQUETAS
            oDoc.WTextBox 172, 280, 12, 450, "Tasa Propuesta(TEM):", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 181, 280, 12, 450, "Tasa Anterior (TEM):", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 190, 280, 12, 450, "Destino de Crédito:", "F3", lnFontSizeBody, hLeft, , , , , , 2
        'DATOS
            oDoc.WTextBox 172, 360, 12, 450, Format(pR!Tasa_Interes, gcFormDato), "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 181, 360, 12, 450, Format(pR!TEMAnterior, gcFormDato), "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 190, 360, 12, 450, pR!cDestino, "F3", lnFontSizeBody, hLeft, , , , , , 2
            'ALPA 20141030 *********************
            If CInt(pR!bExononeracionTasa) = 1 Then
                oDoc.WTextBox 210, 360, 12, 450, "EXONERADO DE TASA", "F3", lnFontSizeBody, hLeft, , , , , , 2
            End If
            'ALPA FIN **************************
    'DERECHA
        'ETIQUETAS
            oDoc.WTextBox 172, 400, 12, 450, "TEA Prop:", "F3", 7, hLeft, , , , , , 2
            oDoc.WTextBox 181, 400, 12, 450, "TEA Ant:", "F3", 7, hLeft, , , , , , 2
        'DATOS
            oDoc.WTextBox 172, 460, 12, 450, Round(((((pR!Tasa_Interes / 100) + 1) ^ (360 / 30)) - 1) * 100, 2), "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 181, 460, 12, 450, Round(((((pR!TEMAnterior / 100) + 1) ^ (360 / 30)) - 1) * 100, 2), "F3", lnFontSizeBody, hLeft, , , , , , 2
    'FIN SECCION Nº 2
    'SECCION Nº 3
        oDoc.WTextBox 206, 55 + 2, 60, 490, "GARANTIAS", "F4", lnFontSizeBody, hLeft
        'CABECERA
            oDoc.WTextBox 215, 57, 12, 36, "COD", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox 215, 93, 12, 60, "Tipo de Garantia", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox 215, 153, 12, 15, "RG", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox 215, 188 - 20, 12, 76, "Documento", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox 215, 264 - 20, 12, 76, "Dirección", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox 215, 340 - 20, 12, 20, "Mon.", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox 215, 367 - 27, 12, 40, "Val. Come.", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox 215, 407 - 27, 12, 40, "Val. Real.", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox 215, 447 - 27, 12, 40, "Val. Utilz.", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox 215, 487 - 27, 12, 40, "Val. Disp.", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox 215, 527 - 27, 12, 40, "Val. Grav.", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox 215, 567 - 27, 12, 25, "PROP.", "F1", nFTablaCabecera, hCenter, , , 1, , , 2

            Dim nCanGarant As Integer

            Dim nValComer As Double
            Dim nValReali As Double
            Dim nValUtiliz As Double
            Dim nValDisp As Double
            Dim nValGrava As Double


            nPosicion = 227

            If Not (prsCredGarant.BOF And prsCredGarant.EOF) Then
                If prsCredGarant.RecordCount > 0 Then
                'psRiesgo = "Riesgo 2" 'Comentado por JOEP 20170420
                
                    For i = 1 To prsCredGarant.RecordCount
                        Dim nAltoAdic As Integer
                        Dim nValorMayor As Integer
                        'RECO20140804****************************
                        Dim nValCom As Double
                        Dim nValRea As Double
                        Dim nValUti As Double
                        Dim nValDis As Double
                        Dim nValGra As Double
                        
                        nValCom = IIf(prsCredGarant!cMoneda = "ME", prsCredGarant!nTasacion * pnTpoCambio, prsCredGarant!nTasacion)
                        nValRea = IIf(prsCredGarant!cMoneda = "ME", prsCredGarant!nRealizacion * pnTpoCambio, prsCredGarant!nRealizacion)
                        nValUti = IIf(prsCredGarant!cMoneda = "ME", prsCredGarant!nGravado * pnTpoCambio, prsCredGarant!nGravado)
                        nValDis = IIf(prsCredGarant!cMoneda = "ME", prsCredGarant!nDisponible * pnTpoCambio, prsCredGarant!nDisponible)
                        nValGra = IIf(prsCredGarant!cMoneda = "ME", prsCredGarant!nValorGravado * pnTpoCambio, prsCredGarant!nValorGravado)
                        'RECO FIN********************************
                        
                        nValorMayor = ValorMayor(Len(prsCredGarant!cTpoGarant), Len(prsCredGarant!cClasGarant), Len(prsCredGarant!cDocDesc), Len(prsCredGarant!cDireccion))

                        nAltoAdic = (nValorMayor / 13) * 6
                        
                        'INICIO Agrego JOEP 20170420
                            If prsCredGarant!nTratLegal = 1 Then
                                psRiesgo = "Riesgo 2"
                                nValorR2 = nValorR2 + 1
                            Else
                                psRiesgo = "Riesgo 1"
                                nValorR1 = nValorR1 + 1
                            End If
                        'FIN Agrego JOEP 20170420
                        
                        'If Trim(prsCredGarant!cClasGarant) = "GARANTIAS NO PREFERIDAS" Then
                            'psRiesgo = "Riesgo 1"
                        'End If
                        
                        oDoc.WTextBox nPosicion + a, 57, 12 + nAltoAdic, 36, prsCredGarant!cNumGarant, "F1", nFTabla, hCenter, , , 1, , , 2
                        oDoc.WTextBox nPosicion + a, 93, 12 + nAltoAdic, 60, prsCredGarant!cTpoGarant, "F1", nFTabla, hjustify, , , 1, , , 2
                        
                        oDoc.WTextBox nPosicion + a, 153, 12 + nAltoAdic, 15, IIf(prsCredGarant!nTratLegal = 1, 2, 1), "F1", nFTabla, hCenter, , , 1, , , 2 'Agrego JOEP 20170420
                        
                        'oDoc.WTextBox nPosicion + A, 153, 12 + nAltoAdic, 15, IIf(Trim(prsCredGarant!cClasGarant) = "GARANTIAS NO PREFERIDAS", 1, 2), "F1", nFTabla, hCenter, , , 1, , , 2
                        
                        oDoc.WTextBox nPosicion + a, 188 - 20, 12 + nAltoAdic, 76, prsCredGarant!cDocDesc & " - Nº " & prsCredGarant!cNroDoc, "F1", nFTabla, hjustify, , , 1, , , 2
                        oDoc.WTextBox nPosicion + a, 264 - 20, 12 + nAltoAdic, 76, prsCredGarant!cDireccion, "F1", nFTabla, hjustify, , , 1, , , 2
                        oDoc.WTextBox nPosicion + a, 340 - 20, 12 + nAltoAdic, 20, prsCredGarant!cMoneda, "F1", nFTabla, hCenter, , , 1, , , 2
                        'RECO20140804*************************************************
                        oDoc.WTextBox nPosicion + a, 367 - 27, 12 + nAltoAdic, 40, Format(nValCom, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
                        oDoc.WTextBox nPosicion + a, 407 - 27, 12 + nAltoAdic, 40, Format(nValRea, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
                        oDoc.WTextBox nPosicion + a, 447 - 27, 12 + nAltoAdic, 40, Format(nValUti, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
                        oDoc.WTextBox nPosicion + a, 487 - 27, 12 + nAltoAdic, 40, Format(nValDis, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
                        oDoc.WTextBox nPosicion + a, 527 - 27, 12 + nAltoAdic, 40, Format(nValGra, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
                        'oDoc.WTextBox nPosicion + a, 367 - 27, 12 + nAltoAdic, 40, Format(prsCredGarant!nTasacion, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
                        'oDoc.WTextBox nPosicion + a, 407 - 27, 12 + nAltoAdic, 40, Format(prsCredGarant!nRealizacion, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
                        'oDoc.WTextBox nPosicion + a, 447 - 27, 12 + nAltoAdic, 40, Format(prsCredGarant!nGravado, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
                        'oDoc.WTextBox nPosicion + a, 487 - 27, 12 + nAltoAdic, 40, Format(prsCredGarant!ndisponible, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
                        'oDoc.WTextBox nPosicion + a, 527 - 27, 12 + nAltoAdic, 40, Format(prsCredGarant!nValorGravado, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
                        'RECO FIN*****************************************************
                        oDoc.WTextBox nPosicion + a, 567 - 27, 12 + nAltoAdic, 25, Mid(prsCredGarant!cRelGarant, 1, 3) & ".", "F1", nFTabla, hCenter, , , 1, , , 2

                        'RECO20140804*********************************
                        nValComer = nValComer + nValCom
                        nValReali = nValReali + nValRea
                        nValUtiliz = nValUtiliz + nValUti
                        nValDisp = nValDisp + nValDis
                        nValGrava = nValGrava + nValGra
                        'nValComer = nValComer + prsCredGarant!nTasacion
                        'nValReali = nValReali + prsCredGarant!nRealizacion
                        'nValUtiliz = nValUtiliz + prsCredGarant!nGravado
                        'nValDisp = nValDisp + prsCredGarant!ndisponible
                        'nValGrava = nValGrava + prsCredGarant!nValorGravado
                        'RECO FIN*************************************
                        prsCredGarant.MoveNext
                        a = a + 12 + nAltoAdic
                    Next
                End If
            End If
            'nPosicion = 207
            oDoc.WTextBox nPosicion + a, 57, 12, 283, "TOTALES", "F1", nFTabla, hCenter, , , 1, , , 2
            oDoc.WTextBox nPosicion + a, 340, 12, 40, Format(nValComer, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
            oDoc.WTextBox nPosicion + a, 380, 12, 40, Format(nValReali, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
            oDoc.WTextBox nPosicion + a, 420, 12, 40, Format(nValUtiliz, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
            oDoc.WTextBox nPosicion + a, 460, 12, 40, Format(nValDisp, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
            oDoc.WTextBox nPosicion + a, 500, 12, 40, Format(nValGrava, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
            oDoc.WTextBox nPosicion + a, 540, 12, 25, "", "F1", nFTabla, hLeft, , , 1, , , 2
            nPosicion = nPosicion + a + 20
            i = 0
            a = 0
    'FIN SECCION Nº 3
    'SECCION Nº 4
    'FIN SECCION Nº 4
    'SECCION Nº 5
        'RECO20140619************************
            Dim nPosAmp As Integer
            oDoc.WTextBox nPosicion + 2, 55, 60, 490, "COBERTURA DE GARANTIA", "F4", lnFontSizeBody, hLeft
            nPosAmp = nPosicion
            nPosicion = nPosicion + 12
            
            oDoc.WTextBox nPosicion, 55, 12, 95, "Cobertura Exp. Este Crédito", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox nPosicion, 150, 12, 95, "Cobertura Exp. Riesgo Único", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox nPosicion, 245, 12, 95, "Tipo de Riesgo", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            nPosicion = nPosicion + 12
            If lnMontoExpEstCred = 0 Then
                oDoc.WTextBox nPosicion + a, 55, 12, 95, Format(0, gcFormView), "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            Else
                oDoc.WTextBox nPosicion + a, 55, 12, 95, Format(nValDisp / lnMontoExpEstCred, gcFormView), "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            End If
            If lnMontoRiesgoUnico = 0 Then
                oDoc.WTextBox nPosicion + a, 150, 12, 95, Format(0, gcFormView), "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            Else
                oDoc.WTextBox nPosicion + a, 150, 12, 95, Format(nValDisp / lnMontoRiesgoUnico, gcFormView), "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            End If
            
            'INICION JOEP 20170420
            If (nValorR1 >= 1) Then
                oDoc.WTextBox nPosicion + a, 245, 12, 95, "RIESGO 1", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            Else
                oDoc.WTextBox nPosicion + a, 245, 12, 95, UCase(psRiesgo), "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            End If
          'FIN JOEP 20170420
            
            'oDoc.WTextBox nPosicion + A, 245, 12, 95, UCase(psRiesgo), "F1", nFTablaCabecera, hCenter, , , 1, , , 2 'Comentado por JOEP 20170420
            
            nPosicion = nPosicion + 25
        'END RECO***************************
            
        If Not (prsCredAmp.BOF And prsCredAmp.EOF) Then
            oDoc.WTextBox nPosAmp + 2, 360, 60, 490, "AMPLIACIÓN DE CRÉDITO", "F4", lnFontSizeBody, hLeft
            
            
            nPosAmp = nPosAmp + 12
            oDoc.WTextBox nPosAmp, 360, 12, 77, "Crédito Nº", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox nPosAmp, 437, 12, 60, "Saldo Capital", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            nPosAmp = nPosAmp + 12
            For i = 1 To prsCredAmp.RecordCount
                oDoc.WTextBox nPosAmp + a, 360, 12, 77, prsCredAmp!cCtaCodAmp, "F1", nFTablaCabecera, hCenter, , , 1, , , 2
                oDoc.WTextBox nPosAmp + a, 437, 12, 60, Format(prsCredAmp!nMonto, gcFormView), "F1", nFTablaCabecera, hRight, , , 1, , , 2
                prsCredAmp.MoveNext
                a = a + 12
            Next
            nPosicion = nPosAmp + a + 10
        End If
    'FIN SECCION Nº 5
    'SECCION Nº 6
        Dim nCanIFs As Integer
        Dim nAdicionaFila As Integer
        If Not (pRRelaBcos.EOF And pRRelaBcos.BOF) Then
            nCanIFs = pRRelaBcos.RecordCount
        End If
        If nCanIFs > 5 Then
            nAdicionaFila = nAdicionaFila + 15
        End If
        nPosicion = nPosicion
        oDoc.WTextBox nPosicion + 6, 55, 60, 490, "RATIOS FINANCIEROS", "F4", lnFontSizeBody, hLeft
        oDoc.WTextBox nPosicion + 16, 55, 56 + nAdicionaFila, 240, "", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 16, 295, 56 + nAdicionaFila, 240, "", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 16, 295, 12, 139, "Institución", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 16, 434, 12, 28, "Moneda", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 16, 462, 12, 40, "Saldo", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 16, 502, 12, 33, "Relacion", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        nPosicion = nPosicion + 9
        oDoc.WTextBox nPosicion + 12, 57, 12, 63, "Liquidez", "F1", 5, hLeft, , , , , , 2
        'ENDEUDADAMIENTO CON OTRAS IFIS**************************************************
           Dim nIndx As Integer
            Dim nTmpPosic As Integer
            nTmpPosic = nPosicion + 8
              
            For nIndx = 1 To pRRelaBcos.RecordCount
                'If Len(pRRelaBcos!Nombre) > 30 Then 'COMENTADO POR APRI 20170719 - MEJORA
                    oDoc.WTextBox nTmpPosic + 9, 295, 9, 139, pRRelaBcos!Nombre, "F1", 5, hLeft, , , , , , 2
                'Else 'COMENTADO POR APRI 20170719 - MEJORA
                    'oDoc.WTextBox nTmpPosic + 9, 295, 9, 139, pRRelaBcos!Nombre, "F1", nFTablaCabecera, hLeft, , , , , , 2 'COMENTADO POR APRI 20170719 - MEJORA
                'End If 'COMENTADO POR APRI 20170719 - MEJORA
                oDoc.WTextBox nTmpPosic + 9, 434, 9, 28, pRRelaBcos!Moneda, "F1", 5, hCenter, , , , , , 2
                oDoc.WTextBox nTmpPosic + 9, 462, 9, 40, Format(pRRelaBcos!Saldo, gcFormView), "F1", 5, hRight, , , , , , 2
                oDoc.WTextBox nTmpPosic + 9, 502, 9, 33, Mid(pRRelaBcos!Relacion, 1, 3) & ".", "F1", 5, hLeft, , , , , , 2
                'pRRelaBcos.MoveNext
                'APRI 20170719 - MEJORA
                If Len(pRRelaBcos!Nombre) > 42 Then
                nTmpPosic = nTmpPosic + 12
                Else
                 nTmpPosic = nTmpPosic + 8
                End If
                pRRelaBcos.MoveNext
                'END APRI
            Next
        '********************************************************************************
'        oDoc.WTextBox nPosicion + 12, 120, 12, 63, Format(lnliquidez, gcFormView), "F1", nFTablaCabecera, hRight, , , 1, , , 2
'        oDoc.WTextBox nPosicion + 12, 188, 12, 63, "Patrimonio", "F1", 5, hLeft, , , , , , 2
'        oDoc.WTextBox nPosicion + 12, 221, 12, 63, Format(lnPatrimonio, gcFormView), "F1", nFTablaCabecera, hRight, , , 1, , , 2
'        nPosicion = nPosicion + 12
'        oDoc.WTextBox nPosicion + 12, 57, 12, 63, "Endeudamiento Patrimonial", "F1", 5, hLeft, , , , , , 2
'        oDoc.WTextBox nPosicion + 12, 120, 12, 63, Format(lnEndeudamiento * 100, "0.00") & "%", "F1", nFTablaCabecera, hRight, , , 1, , , 2
'        oDoc.WTextBox nPosicion + 12, 188, 12, 63, "Inventario", "F1", 5, hLeft, , , , , , 2
'        oDoc.WTextBox nPosicion + 12, 221, 12, 63, Format(lnInventario, gcFormView), "F1", nFTablaCabecera, hRight, , , 1, , , 2
'        nPosicion = nPosicion + 12
'        oDoc.WTextBox nPosicion + 12, 57, 12, 63, "Rentabilidad Patrimonial", "F1", 5, hLeft, , , , , , 2
'        oDoc.WTextBox nPosicion + 12, 120, 12, 63, Format(lnRentabPatrimonial * 100, "0.00") & "%", "F1", nFTablaCabecera, hRight, , , 1, , , 2
'        oDoc.WTextBox nPosicion + 12, 188, 12, 63, "Excedente", "F1", 5, hLeft, , , , , , 2
'        oDoc.WTextBox nPosicion + 12, 221, 12, 63, Format(lnExcedente, gcFormView), "F1", nFTablaCabecera, hRight, , , 1, , , 2
'        nPosicion = nPosicion + 12
'        oDoc.WTextBox nPosicion + 12, 57, 12, 63, "Capacid. De Pago", "F1", 5, hLeft, , , , , , 2
'        oDoc.WTextBox nPosicion + 12, 120, 12, 63, Format(lnCapacidadPago * 100, "0.00") & "%", "F1", nFTablaCabecera, hRight, , , 1, , , 2
        'DATOS RATIOS FINANCIEROS FORMATOS EVALUACION ********************************************************************************
        oDoc.WTextBox nPosicion + 12, 120, 12, 63, IIf(pDRRatiosF!nLiquidezCte = 0, "N/A", Format(pDRRatiosF!nLiquidezCte, gcFormView)), "F1", nFTablaCabecera, hRight, , , 1, , , 2
        oDoc.WTextBox nPosicion + 12, 188, 12, 63, "Patrimonio", "F1", 5, hLeft, , , , , , 2
        oDoc.WTextBox nPosicion + 12, 221, 12, 63, IIf(pDRRatiosF!nPatrimonio = 0, "N/A", Format(pDRRatiosF!nPatrimonio, gcFormView)), "F1", nFTablaCabecera, hRight, , , 1, , , 2
        nPosicion = nPosicion + 12
        oDoc.WTextBox nPosicion + 12, 57, 12, 63, "Endeudamiento Patrimonial", "F1", 5, hLeft, , , , , , 2
        oDoc.WTextBox nPosicion + 12, 120, 12, 63, IIf(pDRRatiosF!nEndeuPat = 0, "N/A", Format(pDRRatiosF!nEndeuPat * 100, "0.00") & "%"), "F1", nFTablaCabecera, hRight, , , 1, , , 2
        oDoc.WTextBox nPosicion + 12, 188, 12, 63, "Ingreso N", "F1", 5, hLeft, , , , , , 2
        oDoc.WTextBox nPosicion + 12, 221, 12, 63, IIf(pDRRatiosF!nIngreNeto = 0, "N/A", Format(pDRRatiosF!nIngreNeto, gcFormView)), "F1", nFTablaCabecera, hRight, , , 1, , , 2
        nPosicion = nPosicion + 12
        oDoc.WTextBox nPosicion + 12, 57, 12, 63, "Rentabilidad Patrimonial", "F1", 5, hLeft, , , , , , 2
        oDoc.WTextBox nPosicion + 12, 120, 12, 63, IIf(pDRRatiosF!nRentaPatri = 0, "N/A", Format(pDRRatiosF!nRentaPatri * 100, "0.00") & "%"), "F1", nFTablaCabecera, hRight, , , 1, , , 2
        oDoc.WTextBox nPosicion + 12, 188, 12, 63, "Excedente", "F1", 5, hLeft, , , , , , 2
        oDoc.WTextBox nPosicion + 12, 221, 12, 63, IIf(pDRRatiosF!nExceMensual, "N/A", Format(pDRRatiosF!nExceMensual, gcFormView)), "F1", nFTablaCabecera, hRight, , , 1, , , 2
        nPosicion = nPosicion + 12
        oDoc.WTextBox nPosicion + 12, 57, 12, 63, "Capacid. De Pago", "F1", 5, hLeft, , , , , , 2
        oDoc.WTextBox nPosicion + 12, 120, 12, 63, IIf(pDRRatiosF!nCapPagNeta = 0, "N/A", Format(pDRRatiosF!nCapPagNeta * 100, "0.00") & "%"), "F1", nFTablaCabecera, hRight, , , 1, , , 2
        'FIN DATOS RATIOS FINANCIEROS FORMATOS EVALUACION *******************************************************************************
        oDoc.WTextBox nPosicion + 12, 188, 12, 63, "Cuota", "F1", 5, hLeft, , , , , , 2
        oDoc.WTextBox nPosicion + 12, 221, 12, 63, Format(lnCuota, gcFormView), "F1", nFTablaCabecera, hRight, , , 1, , , 2
        nPosicion = nPosicion + 12 + 20 + nAdicionaFila 'aqui+ 20
    'FIN SECCION Nº 6
    'SECCION Nº 7
        oDoc.WTextBox nPosicion, 55, 60, 490, "CALIFICACIÓN Y RELACIÓN DE TITULARES / CÓNYUGE / AVALES", "F4", 7, hLeft
        oDoc.WTextBox nPosicion + 12, 55, 56, 480, "", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 12, 55, 12, 240, "Nombre", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 12, 295, 12, 115, "Relación", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 12, 410, 12, 25, "Normal", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 12, 435, 12, 25, "Poten.", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 12, 460, 12, 25, "Defic.", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 12, 485, 12, 25, "Dudos", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 12, 510, 12, 25, "Pérdida", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        'CALIFICACION Y RELACION DE TITULARES/CONYUGUE /AVALES**************************************************
        If Not (prsCalfSBSRela.EOF And prsCalfSBSRela.BOF) Then
            Dim nIndx2 As Integer
            Dim nTmpPosic2 As Integer
            nTmpPosic2 = nPosicion + 14
            For nIndx2 = 1 To prsCalfSBSRela.RecordCount
                oDoc.WTextBox nTmpPosic2 + 9, 55, 11, 240, prsCalfSBSRela!cPersNombre, "F1", nFTablaCabecera, hLeft, , , , , , 1
                oDoc.WTextBox nTmpPosic2 + 9, 295, 11, 115, prsCalfSBSRela!cConsDescripcion, "F1", nFTablaCabecera, hLeft, , , , , , 1
                oDoc.WTextBox nTmpPosic2 + 9, 410, 11, 25, prsCalfSBSRela!Normal & "%", "F1", nFTablaCabecera, hCenter, , , , , , 1
                oDoc.WTextBox nTmpPosic2 + 9, 435, 11, 25, prsCalfSBSRela!Potencial & "%", "F1", nFTablaCabecera, hCenter, , , , , , 1
                oDoc.WTextBox nTmpPosic2 + 9, 460, 11, 25, prsCalfSBSRela!DEFICIENTE & "%", "F1", nFTablaCabecera, hCenter, , , , , , 1
                oDoc.WTextBox nTmpPosic2 + 9, 485, 11, 25, prsCalfSBSRela!DUDOSO & "%", "F1", nFTablaCabecera, hCenter, , , , , , 1
                oDoc.WTextBox nTmpPosic2 + 9, 510, 11, 25, prsCalfSBSRela!PERDIDA & "%", "F1", nFTablaCabecera, hCenter, , , , , , 1
                prsCalfSBSRela.MoveNext
                nTmpPosic2 = nTmpPosic2 + 8
            Next
        End If
    'FIN SECCION Nº 7
     'APRI20170719 TI-ERS025-2017
    If prsPersVinc.RecordCount > 0 Then
        nPosicion = nPosicion + 70
        oDoc.WTextBox nPosicion, 55, 60, 490, "CALIFICACIÓN DE VINCULADOS AL COLABORADOR", "F4", 7, hLeft
        oDoc.WTextBox nPosicion + 9, 55, 56, 480, "", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 9, 55, 12, 240, "Nombre", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 9, 295, 12, 115, "Relación", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 9, 410, 12, 25, "Normal", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 9, 435, 12, 25, "Poten.", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 9, 460, 12, 25, "Defic.", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 9, 485, 12, 25, "Dudos", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 9, 510, 12, 25, "Pérdida", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        If Not (prsPersVinc.EOF And prsPersVinc.BOF) Then
            Dim nIndxP As Integer
            Dim nTmpPosicP As Integer
            nTmpPosicP = nPosicion + 14
            For nIndxP = 1 To prsPersVinc.RecordCount
                oDoc.WTextBox nTmpPosicP + 9, 55, 11, 240, prsPersVinc!cPersNombre, "F1", nFTablaCabecera, hLeft, , , , , , 1
                oDoc.WTextBox nTmpPosicP + 9, 295, 11, 115, prsPersVinc!cRelacion, "F1", nFTablaCabecera, hLeft, , , , , , 1
                oDoc.WTextBox nTmpPosicP + 9, 410, 11, 28, prsPersVinc!Normal & "%", "F1", nFTablaCabecera, hCenter, , , , , , 1
                oDoc.WTextBox nTmpPosicP + 9, 435, 11, 28, prsPersVinc!Potencial & "%", "F1", nFTablaCabecera, hCenter, , , , , , 1
                oDoc.WTextBox nTmpPosicP + 9, 460, 11, 28, prsPersVinc!DEFICIENTE & "%", "F1", nFTablaCabecera, hCenter, , , , , , 1
                oDoc.WTextBox nTmpPosicP + 9, 485, 11, 28, prsPersVinc!DUDOSO & "%", "F1", nFTablaCabecera, hCenter, , , , , , 1
                oDoc.WTextBox nTmpPosicP + 9, 510, 11, 28, prsPersVinc!PERDIDA & "%", "F1", nFTablaCabecera, hCenter, , , , , , 1
                prsPersVinc.MoveNext
                nTmpPosicP = nTmpPosicP + 8
            Next
        End If
    End If
    'END APRI
    'WIOR 20160621 *** SOBREENDEUDAMIENTO DE CLIENTES
    nPosicion = nPosicion + 70
    If Not (prsSobreEnd.EOF And prsSobreEnd.BOF) Then
        oDoc.WTextBox nPosicion, 55, 56, 330, "SOBREENDEUDAMIENTO DE CLIENTES", "F4", 7, hLeft
        nPosicion = nPosicion + 12
        oDoc.WTextBox nPosicion, 55, 12, 80, "Deuda Potencial", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion, 135, 12, 80, Format(CCur(prsSobreEnd!nDeudaTotal), "###," & String(15, "#") & "#0.00"), "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        If Not (prsSobreEndCodigos.EOF And prsSobreEndCodigos.BOF) Then
            nPosicion = nPosicion + 18

            oDoc.WTextBox nPosicion, 55, 12, 50, "Códigos", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox nPosicion, 105, 12, 110, "Resultados", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox nPosicion, 215, 12, 160, "Detalle", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox nPosicion, 375, 12, 160, "Plan de Mitigación del Riesgo", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            nPosicion = nPosicion + 12

            For i = 1 To prsSobreEndCodigos.RecordCount
                oDoc.WTextBox nPosicion, 55, 15, 50, Trim(prsSobreEndCodigos!cCodigo), "F1", nFTablaCabecera, hCenter, vMiddle, , 1, , , 2
                oDoc.WTextBox nPosicion, 105, 15, 110, Trim(prsSobreEndCodigos!cResultado), "F1", nFTablaCabecera, hLeft, vMiddle, , 1, , , 2
                oDoc.WTextBox nPosicion, 215, 15, 160, Trim(prsSobreEndCodigos!cDetalle), "F1", 6, hLeft, vMiddle, , 1, , , 2
                oDoc.WTextBox nPosicion, 375, 15, 160, Trim(prsSobreEndCodigos!cPlanmitigacion), "F1", 6, hLeft, vMiddle, , 1, , , 2
                prsSobreEndCodigos.MoveNext
                nPosicion = nPosicion + 15
            Next i
        End If
        nPosicion = nPosicion + 12
    End If
    'WIOR FIN ********
    
    'RECO NUEVO
    'nPosicion = nPosicion + 70
    If Not (prsComentAnalis.BOF And prsComentAnalis.EOF) Then
        oDoc.WTextBox nPosicion, 55, 56, 330, "COMENTARIO ANALISTA", "F1", nFTablaCabecera, hLeft
        nPosicion = nPosicion + 9
        oDoc.WTextBox nPosicion, 55, 40, 480, prsComentAnalis!cComentAnalista, "F1", nFTablaCabecera, hjustify, , , 1, , , 2
        nPosicion = nPosicion + 48
    End If
    
    'RECO FIN NUEVO
    'SECCION Nº 8
        'oDoc.WTextBox nPosicion, 55, 56, 280, "", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion, 55, 100, 210, "", "F1", nFTablaCabecera, hCenter, , , 1, , , 2 'FRHU 20160811 Anexo002
        'oDoc.WTextBox nPosicion, 55, 56, 165, "", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        'oDoc.WTextBox nPosicion, 55, 12, 165, "EXONERACIONES", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        'oDoc.WTextBox nPosicion, 55, 12, 165, "AUTORIZACIONES", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        'oDoc.WTextBox nPosicion, 55, 12, 280, "AUTORIZACIONES", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion, 55, 12, 210, "Autorizaciones", "F1", nFTablaCabecera, hCenter, , , 1, , , 2 'FRHU 20160811 Anexo002
        'oDoc.WTextBox nPosicion, 385, 12, 200, "NIVELES DE APROBACION POR EXPOSICION", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion, 265, 12, 270, "Nivel de Aprobación", "F1", nFTablaCabecera, hCenter, , , 1, , , 2 'FRHU 20160811 Anexo002 ERS002-2016
        'oDoc.WTextBox nPosicion, 385, 56, 200, "", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion, 265, 100, 270, "", "F1", nFTablaCabecera, hCenter, , , 1, , , 2 'FRHU 20160811 Anexo002
        
        nPosicion = nPosicion + 12 'FRHU 20160811 Anexo002 ERS002-2016 NEW
        oDoc.WTextBox nPosicion, 265, 12, 135, "Por Autorización", "F1", nFTablaCabecera, hCenter, , , 1, , , 2 'FRHU 20160811 Anexo002 ERS002-2016 NEW
        oDoc.WTextBox nPosicion, 400, 12, 135, "Por Exposición", "F1", nFTablaCabecera, hCenter, , , 1, , , 2 'FRHU 20160811 Anexo002 ERS002-2016 NEW
        oDoc.WTextBox nPosicion, 400, 88, 135, "", "F1", nFTablaCabecera, hCenter, , , 1, , , 2 'FRHU 20160811 Anexo002 ERS002-2016 NEW
        
        If psHabilitaNiveles = 0 Then
            If Not (prsExoAutCred.EOF And prsExoAutCred.BOF) Then
                Dim nIndx3 As Integer
                Dim nTmpPosic3 As Integer
                Dim nTmpPosicExo As Integer
                Dim nTmpPosicAut As Integer
                nTmpPosic3 = nPosicion + 14
                nTmpPosicExo = nPosicion + 2
                nTmpPosicAut = nPosicion + 2
                'For nIndx3 = 1 To 3
                For nIndx3 = 1 To prsExoAutCred.RecordCount
                 Dim texto As String
                    If prsExoAutCred!nTipoExoneraCod = 1 Then
                        Set lrsRespNvlExo = oDCredExoAut.RecuperaRespExoAut(prsExoAutCred!cExoneraCod, IIf(lnMontoRiesgoUnico = 0, lnMontoExpEstCred, lnMontoRiesgoUnico))
                        oDoc.WTextBox nTmpPosicExo + 9, 55, 9, 150, prsExoAutCred!cExoneraDesc, "F1", nFTablaCabecera, hLeft, , , , , , 2
                        
                        If Not (lrsRespNvlExo.BOF And lrsRespNvlExo.EOF) Then
                            oDoc.WTextBox nTmpPosicExo + 9, 145, 9, 150, lrsRespNvlExo!cNivAprDesc, "F1", nFTablaCabecera, hLeft, , , , , , 2
                        End If
                        nTmpPosicExo = nTmpPosicExo + 5
                    Else
                        Set lrsRespNvlExo = oDCredExoAut.RecuperaRespExoAut(prsExoAutCred!cExoneraCod, IIf(lnMontoRiesgoUnico = 0, lnMontoExpEstCred, lnMontoRiesgoUnico))
                        oDoc.WTextBox nTmpPosicAut + 9, 225, 9, 240, prsExoAutCred!cExoneraDesc, "F1", nFTablaCabecera, hLeft, , , , , , 2
                        If Not (lrsRespNvlExo.BOF And lrsRespNvlExo.EOF) Then
                            oDoc.WTextBox nTmpPosicAut + 9, 295, 9, 150, lrsRespNvlExo!cNivAprDesc, "F1", nFTablaCabecera, hLeft, , , , , , 2
                        End If
                        nTmpPosicAut = nTmpPosicAut + 5
                    End If
                    prsExoAutCred.MoveNext
                    Set lrsRespNvlExo = Nothing
                Next
            End If
            
            If Not (prsCredResulNivApr.EOF And prsCredResulNivApr.BOF) Then
                Dim nIndx4 As Integer
                Dim nTmpPosic4 As Integer
    
                nTmpPosic4 = nPosicion + 14
                For nIndx = 1 To prsCredResulNivApr.RecordCount
                    oDoc.WTextBox nTmpPosic4 + 12, 390, 12, 305, prsCredResulNivApr!cNivAprDesc, "F1", nFTablaCabecera, hLeft
                    nTmpPosic4 = nTmpPosic4 + 5
                    prsCredResulNivApr.MoveNext
                Next
            End If
        End If
        'FRHU 20160811 Anexo002 ERS002-2016
        If Not (prsAutorizaciones.EOF And prsAutorizaciones.BOF) Then
            Dim nTmpPosicExo5 As Integer
            nTmpPosicExo5 = nPosicion + 2
            For nIndx = 1 To prsAutorizaciones.RecordCount
                If Len(prsAutorizaciones!cExoneraDesc) <= 69 Then
                    oDoc.WTextBox nTmpPosicExo5 + 9, 55, 9, 280, prsAutorizaciones!cExoneraDesc, "F1", nFTablaCabecera, hLeft, , , , , , 2
                    oDoc.WTextBox nTmpPosicExo5 + 11, 267, 12, 305, prsAutorizaciones!cNivAprDesc, "F1", nFTablaCabecera, hLeft
                    nTmpPosicExo5 = nTmpPosicExo5 + 9
                Else
                    oDoc.WTextBox nTmpPosicExo5 + 9, 55, 9, 280, prsAutorizaciones!cExoneraDesc, "F1", nFTablaCabecera, hLeft, , , , , , 2
                    oDoc.WTextBox nTmpPosicExo5 + 11, 267, 12, 305, prsAutorizaciones!cNivAprDesc, "F1", nFTablaCabecera, hLeft
                    nTmpPosicExo5 = nTmpPosicExo5 + 18
                End If
                prsAutorizaciones.MoveNext
            Next nIndx
        End If
        'FIN FRHU 20160811
    'FIN SECCION Nº 8
    'SECCION Nº9
    'nPosicion = nPosicion + 50
    nPosicion = nPosicion + 100 'FRHU 20160811 ERS002-2016
    oDoc.WTextBox nPosicion + 15, 150, 56, 150, "RESOLUCION DE COMITÉ, EN CONCLUSION: ", "F1", nFTablaCabecera, Left
    nPosicion = nPosicion + 12
    oDoc.WTextBox nPosicion + 15, 150, 56, 70, "MONTO", "F1", nFTablaCabecera, Left
    oDoc.WTextBox nPosicion + 15, 250, 56, 70, "CUOTAS", "F1", nFTablaCabecera, Left
    oDoc.WTextBox nPosicion + 15, 350, 56, 70, "TI", "F1", nFTablaCabecera, Left
    oDoc.WTextBox nPosicion + 15, 450, 56, 70, "VCTO", "F1", nFTablaCabecera, Left
    
    nPosicion = nPosicion + 20
    oDoc.WTextBox nPosicion + 15, 70, 56, 150, "APROBADO POR: ", "F1", nFTablaCabecera, Left
    oDoc.WTextBox nPosicion + 15, 150, 56, 70, "...................", "F1", nFTablaCabecera, Left
    oDoc.WTextBox nPosicion + 15, 250, 56, 70, "...................", "F1", nFTablaCabecera, Left
    oDoc.WTextBox nPosicion + 15, 350, 56, 70, "...................", "F1", nFTablaCabecera, Left
    oDoc.WTextBox nPosicion + 15, 450, 56, 70, "...................", "F1", nFTablaCabecera, Left
    'FIN SECCION N9
    oDoc.PDFClose
    oDoc.Show
    Exit Sub
ErrorImprimirPDF:
    MsgBox Err.Description, vbInformation, "Aviso"
End Sub

Private Sub ImprimeHojaAprobacionCredConsumo(ByRef pR As ADODB.Recordset, ByRef pRB As ADODB.Recordset, ByRef pRCliRela As ADODB.Recordset, ByRef pRRelaBcos As ADODB.Recordset, _
                                      ByRef pRDatFinan As ADODB.Recordset, ByRef pRResGarTitAva As ADODB.Recordset, ByRef pRGarantCred As ADODB.Recordset, _
                                      ByVal psCodCta, ByVal prsCredEval As ADODB.Recordset, ByVal prsCredAmp As ADODB.Recordset, ByVal prsCalfSBSRela As ADODB.Recordset, _
                                      ByVal prsExoAutCred As ADODB.Recordset, ByVal prsRiesgoUnico As ADODB.Recordset, ByVal prsCredGarant As ADODB.Recordset, _
                                      ByVal prsCredResulNivApr As ADODB.Recordset, ByVal REstCivConvenio As ADODB.Recordset, ByVal pRNivApr As ADODB.Recordset, ByVal psRiesgo As String, _
                                      ByVal psHabilitaNiveles As String, ByVal prsOpRiesgo As ADODB.Recordset, ByVal prsComentAnalis As ADODB.Recordset, _
                                      ByVal pnTpoCambio As Double, ByVal prsSobreEnd As ADODB.Recordset, ByVal prsSobreEndCodigos As ADODB.Recordset, _
                                      ByVal pDRRatiosF As ADODB.Recordset, ByVal prsAutorizaciones As ADODB.Recordset, ByVal prsPersVinc As ADODB.Recordset) 'RECO20140804 - Se añadió el parametro pnTpoCambio
                                      'WIOR 20160621 AGREGO ByVal prsSobreEnd As ADODB.Recordset, ByVal prsSobreEndCodigos As ADODB.Recordset
                                      'RECO20160730 SE AGREGO LOS RATIOS DE FORMATOS EVALUACION
                                      'FRHU 20160811 Anexo002 ERS002-2016: Se agrego ByVal prsAutorizaciones As ADODB.Recordset
                                      'APRI 20170719 AGREGO prsPersVinc TI-ERS-025 2017
    Dim oFun As New COMFunciones.FCOMImpresion
    Dim lnliquidez As Double, lnCapacidadPago As Double, lnExcedente As Double
    Dim lnPatriEmpre As Double, lnPatrimonio As Double, lnIngresoNeto As Double
    Dim lnRentabPatrimonial As Double, lnEndeudamiento As Double
    Dim lnInventario As Double, lnCuota As Double
    Dim lnMontoRiesgoUnico As Double
    Dim lnMontoExpEstCred As Double, lnRemNeta As Double, lnEgresos As Double
    'ALPA20160714****************************
    Dim lnCapacidadPago2 As Double
    '****************************************
    On Error GoTo ErrorImprimirPDF
    Dim oDoc  As cPDF
    Dim nTipo As Integer
    Set oDoc = New cPDF
    
    
    Dim i As Integer
    Dim a As Integer
    Dim nPosicion As Integer
    
    
    Dim lrsIngGas As ADODB.Recordset
    Dim rsConDPer As COMDPersona.DCOMPersonas
    Dim oCrDoc As New COMDCredito.DCOMCredDoc
    Dim oDCredExoAut As New COMDCredito.DCOMNivelAprobacion
    Dim lrsRespNvlExo As ADODB.Recordset
    
    Set oDCredExoAut = New COMDCredito.DCOMNivelAprobacion
    
    Dim lrDatosRatioFI As ADODB.Recordset
        
    'CTI3 (ferimoro)
    Dim oDpExposicion As COMDCredito.DCOMCredito 'CTI3 (ferimoro) ERS062-2018
    Dim expo As ADODB.Recordset                 'CTI3 (ferimoro) ERS062-2018
    Dim nRunico As Integer
    Dim nClePref As Integer
    
     Dim oRecCtaRef As COMDCredito.DCOMCredito
     Dim cCtaRefe As ADODB.Recordset
    '**************
    
    Dim dFecFteIng As Date
    
    Dim nIngNeto As Double
    Dim nGasFamiliar  As Double
    Dim nExpoCredAct As Double
    
    'Dim oCliPre As New COMNCredito.NCOMCredito 'COMENTADO POR ARLO 20170722
    Dim bValidarCliPre As Boolean
     
    If oCrDoc.ObtenerFechaProNumFuente(pR!cNumeroFuente, gdFecSis).RecordCount > 0 Then
        dFecFteIng = Format(oCrDoc.ObtenerFechaProNumFuente(pR!cNumeroFuente, gdFecSis)!valor, "dd/mm/yyyy")
    End If
    
    Set lrDatosRatioFI = New ADODB.Recordset
    
    Set lrDatosRatioFI = oCrDoc.ObtieneRatiosFI(pR!cNumeroFuente, psCodCta)
    lnliquidez = 0
    lnEndeudamiento = 0
    lnRentabPatrimonial = 0
    lnCapacidadPago = 0
    lnPatrimonio = 0
    lnInventario = 0
    lnExcedente = 0
    lnCuota = 0
    lnRemNeta = 0
    lnEgresos = 0
        
    nValorR1 = 0 'JOEP20170627
    nValorR2 = 0 'JOEP20170627
    
    If Not (lrDatosRatioFI.EOF And lrDatosRatioFI.BOF) Then
        Dim nIndexRatios As Integer
        For nIndexRatios = 0 To lrDatosRatioFI.RecordCount - 1
            If lrDatosRatioFI!nValor = 1 Then
                'lnInventario = lrDatosRatioFI!nMonto
            ElseIf lrDatosRatioFI!nValor = 2 Then
                'lnExcedente = lrDatosRatioFI!nMonto
            ElseIf lrDatosRatioFI!nValor = 3 Then
                lnCuota = lrDatosRatioFI!nMonto
            ElseIf lrDatosRatioFI!nValor = 4 Then
                'lnPatrimonio = lrDatosRatioFI!nMonto
            ElseIf lrDatosRatioFI!nValor = 5 Then
                'lnRentabPatrimonial = lrDatosRatioFI!nMonto
            ElseIf lrDatosRatioFI!nValor = 9 Then
                lnRemNeta = lrDatosRatioFI!nMonto
            ElseIf lrDatosRatioFI!nValor = 10 Then
                lnEgresos = lrDatosRatioFI!nMonto
              'ALPA20160714*****************************************************************************
            ElseIf lrDatosRatioFI!nValor = 11 Then
                lnCapacidadPago2 = IIf(IsNull(lrDatosRatioFI!nMonto), 0, lrDatosRatioFI!nMonto)
            '*****************************************************************************************
            End If
            lrDatosRatioFI.MoveNext
        Next
        lnExcedente = lnRemNeta - lnEgresos
        'ALPA20160714******************************************************************
        'lnCapacidadPago = lnCuota / IIf(lnExcedente = 0, 1, lnExcedente)
        If lnCapacidadPago2 = 0 Then
            lnCapacidadPago = lnCuota / IIf(lnExcedente = 0, 1, lnExcedente)
        Else
            lnCapacidadPago = lnCapacidadPago2
        End If
        '******************************************************************************
        
    End If
    Set rsConDPer = New COMDPersona.DCOMPersonas
    Set lrsIngGas = rsConDPer.ObtenerDatosHojEvaluaci(pR!cNumeroFuente, dFecFteIng)
    
    
    
    oDoc.Author = gsCodUser
    oDoc.Creator = "SICMACT - Negocio"
    oDoc.Producer = "Caja Municipal de Ahorros y Crédito de Maynas S.A."
    oDoc.Subject = "Hoja de Aprobación de Créditos Nº " & "RECO"
    oDoc.Title = "Hoja de Aprobación de Créditos Nº " & "RECO"
    If Not oDoc.PDFCreate(App.Path & "\Spooler\" & IIf(nTipo = 1, "Previo", "") & psCodCta & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
        Exit Sub
    End If
    
    'FUENTES
    Dim nFTabla As Integer
    Dim nFTablaCabecera As Integer
    Dim lnFontSizeBody As Integer
    
    oDoc.Fonts.Add "F1", "Times New Roman", TrueType, Normal, WinAnsiEncoding 'RECO20140227
    oDoc.Fonts.Add "F2", "Times New Roman", TrueType, Bold, WinAnsiEncoding 'RECO20140227
    oDoc.Fonts.Add "F3", "Arial Narrow", TrueType, Normal, WinAnsiEncoding 'RECO20140227
    oDoc.Fonts.Add "F4", "Arial Narrow", TrueType, Bold, WinAnsiEncoding 'RECO20140227
    
    nFTablaCabecera = 7
    nFTabla = 7
    lnFontSizeBody = 7
    'FIN FUENTES
    
    oDoc.NewPage A4_Vertical

    lnMontoRiesgoUnico = 0#
    
    If Not (prsRiesgoUnico.EOF And prsRiesgoUnico.BOF) Then
        lnMontoRiesgoUnico = prsRiesgoUnico!nMonto
    End If
    
          If Not (lrsIngGas.BOF And lrsIngGas.EOF) Then

            Dim j As Integer
            
            For j = 1 To lrsIngGas.RecordCount
                If (lrsIngGas!cCodHojEval = 300102) Then
                    nIngNeto = Format(lrsIngGas!nUnico, gcFormView)
                ElseIf (lrsIngGas!cCodHojEval = 300201) Then
                    nGasFamiliar = Format(lrsIngGas!nUnico, gcFormView)
                ElseIf (lrsIngGas!cCodHojEval = 400205) Then
                    nExpoCredAct = Format(lrsIngGas!nUnico, gcFormView)
                End If
                lrsIngGas.MoveNext
            Next
         End If
        
            'Set oCliPre = New COMNCredito.NCOMCredito  'COMENTADO POR ARLO 20170722
            'bValidarCliPre = oCliPre.ValidarClientePreferencial(pR!cPersCod) 'COMENTADO POR ARLO 20170722
            bValidarCliPre = False 'ARLO 20170722
        If bValidarCliPre Then
            
        End If
    
    'oDoc.WTextBox 40, 70, 10, 450, "HOJA DE APROBACIÓN", "F3", 11, hCenter, , vbBlack
    'RECO20141011 ***********************************
    If bValidarCliPre = True Then
        nClePref = 2
        oDoc.WTextBox 40, 70, 10, 450, "HOJA DE APROBACIÓN - CLIENTE PREFERENCIAL", "F3", 11, hCenter, , vbBlack
    Else
        nClePref = 1
        oDoc.WTextBox 40, 70, 10, 450, "HOJA DE APROBACIÓN", "F3", 11, hCenter, , vbBlack
    End If
    'RECO END ***************************************
    oDoc.WTextBox 60, 55, 10, 200, "APROBACION DE CREDITOS", "F4", 9, hLeft
    oDoc.WTextBox 60, 354, 10, 200, "FECHA APROBACION:....................", "F4", 9, hRight
    
    'SECCION Nº 1
        oDoc.WTextBox 80, 55, 96, 510, "", "F3", 7, hLeft, , , 1, vbBlack
    'IZQUIERDA
        'ETIQUETAS
            oDoc.WTextBox 80, 57, 12, 450, "Cliente:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 88, 57, 12, 450, "DNI/RUC:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 96, 57, 12, 450, "Edad y Estado Civil", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 104, 57, 12, 450, "Tipo de Credito:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 112, 57, 12, 450, "Producto de Crédito:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 120, 57, 12, 450, "Exposición Anterior Máxima:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 128, 57, 12, 450, "Exposición con este Crédito:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 136, 57, 12, 450, "Exposición Riesgo Único:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 144, 57, 12, 450, "Fecha de Solicitud:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 152, 57, 12, 450, "Campaña:", "F3", lnFontSizeBody, hLeft, , , , , , 2
        'DATOS
            oDoc.WTextBox 80, 145, 12, 450, UCase(pR!Prestatario), "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 88, 145, 12, 450, pR!DniRuc & "/" & pR!Ruc, "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 96, 145, 12, 450, REstCivConvenio!nEdad & "-" & REstCivConvenio!cestciv, "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 104, 145, 12, 450, pR!Tipo_Cred, "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 112, 145, 12, 450, pR!Tipo_Prod, "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 120, 145, 12, 450, Format(pR!ExpoAntMax, gcFormView), "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 128, 145, 12, 450, Format(pR!nMontoExpCredito, gcFormView), "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 136, 145, 12, 450, Format(lnMontoRiesgoUnico, gcFormView), "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 144, 145, 12, 450, pR!Fec_Soli, "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 152, 145, 12, 450, pR!cCampania, "F3", lnFontSizeBody, hLeft, , , , , , 2
            lnMontoExpEstCred = pR!nMontoExpCredito
            
    'DERECHA
        'ETIQUETAS
            oDoc.WTextBox 80, 330, 12, 450, "Nº Crédito:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 88, 330, 12, 450, "Modalidad", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 96, 330, 12, 450, "Cod. Analista:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 104, 330, 12, 450, "Agencia:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 112, 330, 12, 450, "Línea de Crédito:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 120, 330, 12, 450, "Inst. Conv.:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 128, 330, 12, 450, "Antigüedad cliente:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 136, 330, 12, 450, "Giro/Act Negocio:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 144, 330, 12, 450, "Direc. Principal:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 152, 330, 12, 450, "Direc. Negocio:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 160, 330, 12, 450, "Opinión Riesgo:", "F3", lnFontSizeBody, hLeft, , , , , , 2
        'DATOS
            oDoc.WTextBox 80, 400, 12, 450, pR!Nro_Credito, "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 88, 400, 12, 450, UCase(pR!Modalidad), "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 96, 400, 12, 450, UCase(pR!Analista), "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 104, 400, 12, 450, UCase(pR!Oficina), "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 112, 400, 12, 450, pR!Linea_Cred, "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 120, 400, 12, 450, REstCivConvenio!cInstitucion, "F3", lnFontSizeBody, hLeft, , , , , , 2
            'oDoc.WTextBox 128, 400, 12, 450, REstCivConvenio!nAntiguedad & " año ", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 128, 400, 12, 450, REstCivConvenio!nAntiguedad & " años ", "F3", lnFontSizeBody, hLeft, , , , , , 2 'FRHU 20160823
            oDoc.WTextBox 136, 400, 12, 450, pR!ActiGiro, "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 144, 400, 12, 450, pR!dire_domicilio, "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 152, 400, 12, 450, pR!dire_trabajo, "F3", lnFontSizeBody, hLeft, , , , , , 2
            If Not (prsOpRiesgo.EOF And prsOpRiesgo.BOF) Then
                oDoc.WTextBox 160, 400, 12, 450, prsOpRiesgo!cRiesgoValor, "F3", lnFontSizeBody, hLeft, , , , , , 2
            End If
    'FIN SECCION Nº 1
    'SECCION Nº 2
        oDoc.WTextBox 179, 55, 60, 490, "DATOS DEL CRÉDITO", "F4", lnFontSizeBody, hLeft
        oDoc.WTextBox 188, 55, 30, 510, "", "F3", 7, hLeft, , , 1, vbBlack
    'IZQUIERDA
        'ETIQUETAS
            oDoc.WTextBox 188, 57, 12, 450, "Monto Propuesto:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 197, 57, 12, 450, "Plazo de Gracia:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 206, 57, 12, 450, "Plazo de Crédito:", "F3", lnFontSizeBody, hLeft, , , , , , 2
        'DATOS
            oDoc.WTextBox 188, 140, 12, 450, Format(pR!Ptmo_Propto, gcFormView), "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 197, 140, 12, 450, pR!PeriodoGracia, "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 206, 140, 12, 450, pR!Plazo, "F3", lnFontSizeBody, hLeft, , , , , , 2
    'CENTRO
        'ETIQUETAS
            oDoc.WTextBox 188, 280, 12, 450, "Tasa Propuesta(TEM):", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 197, 280, 12, 450, "Tasa Anterior (TEM):", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 206, 280, 12, 450, "Destino de Crédito:", "F3", lnFontSizeBody, hLeft, , , , , , 2
        'DATOS
            oDoc.WTextBox 188, 360, 12, 450, Format(pR!Tasa_Interes, gcFormDato), "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 197, 360, 12, 450, Format(pR!TEMAnterior, gcFormDato), "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 206, 360, 12, 450, pR!cDestino, "F3", lnFontSizeBody, hLeft, , , , , , 2
            'ALPA 20141030 *********************
            If CInt(pR!bExononeracionTasa) = 1 Then
                oDoc.WTextBox 226, 360, 12, 450, "EXONERADO DE TASA", "F3", lnFontSizeBody, hLeft, , , , , , 2
            End If
            'ALPA FIN **************************
    'DERECHA
        'ETIQUETAS
            oDoc.WTextBox 188, 400, 12, 450, "TEA Prop:", "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 197, 400, 12, 450, "TEA Ant:", "F3", lnFontSizeBody, hLeft, , , , , , 2
        'DATOS
            oDoc.WTextBox 188, 460, 12, 450, Round(((((pR!Tasa_Interes / 100) + 1) ^ (360 / 30)) - 1) * 100, 2), "F3", lnFontSizeBody, hLeft, , , , , , 2
            oDoc.WTextBox 197, 460, 12, 450, Round(((((pR!TEMAnterior / 100) + 1) ^ (360 / 30)) - 1) * 100, 2), "F3", lnFontSizeBody, hLeft, , , , , , 2
    'FIN SECCION Nº 2
    'SECCION Nº 3
        oDoc.WTextBox 222, 55 + 2, 60, 490, "GARANTIAS", "F4", lnFontSizeBody, hLeft
        'CABECERA
            oDoc.WTextBox 231, 57, 12, 36, "COD", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox 231, 93, 12, 60, "Tipo de Garantia", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox 231, 153, 12, 15, "RG", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox 231, 188 - 20, 12, 76, "Documento", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox 231, 264 - 20, 12, 76, "Dirección", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox 231, 340 - 20, 12, 20, "Mon.", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox 231, 367 - 27, 12, 40, "Val. Come.", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox 231, 407 - 27, 12, 40, "Val. Real.", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox 231, 447 - 27, 12, 40, "Val. Utilz.", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox 231, 487 - 27, 12, 40, "Val. Disp.", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox 231, 527 - 27, 12, 40, "Val. Grav.", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox 231, 567 - 27, 12, 25, "PROP.", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            
            Dim nCanGarant As Integer
            

            Dim nValComer As Double
            Dim nValReali As Double
            Dim nValUtiliz As Double
            Dim nValDisp As Double
            Dim nValGrava As Double
            
           
            nPosicion = 243
            
            If Not (prsCredGarant.BOF And prsCredGarant.EOF) Then
                If prsCredGarant.RecordCount > 0 Then
                'psRiesgo = "Riesgo 2" 'Comentado por JOEP 20170420
                
                    For i = 1 To prsCredGarant.RecordCount
                        Dim nAltoAdic As Integer
                        Dim nValorMayor As Integer
                        'RECO20140804****************************
                        Dim nValCom As Double
                        Dim nValRea As Double
                        Dim nValUti As Double
                        Dim nValDis As Double
                        Dim nValGra As Double
                        
                        nValCom = IIf(prsCredGarant!cMoneda = "ME", prsCredGarant!nTasacion * pnTpoCambio, prsCredGarant!nTasacion)
                        nValRea = IIf(prsCredGarant!cMoneda = "ME", prsCredGarant!nRealizacion * pnTpoCambio, prsCredGarant!nRealizacion)
                        nValUti = IIf(prsCredGarant!cMoneda = "ME", prsCredGarant!nGravado * pnTpoCambio, prsCredGarant!nGravado)
                        nValDis = IIf(prsCredGarant!cMoneda = "ME", prsCredGarant!nDisponible * pnTpoCambio, prsCredGarant!nDisponible)
                        nValGra = IIf(prsCredGarant!cMoneda = "ME", prsCredGarant!nValorGravado * pnTpoCambio, prsCredGarant!nValorGravado)
                        'RECO FIN********************************
                        nValorMayor = ValorMayor(Len(prsCredGarant!cTpoGarant), Len(prsCredGarant!cClasGarant), Len(prsCredGarant!cDocDesc), Len(prsCredGarant!cDireccion))

                        nAltoAdic = (nValorMayor / 13) * 6
                        
                        'INICIO Agrego JOEP 20170420
                            If prsCredGarant!nTratLegal = 1 Then
                                psRiesgo = "Riesgo 2"
                                nValorR2 = nValorR2 + 1
                            Else
                                psRiesgo = "Riesgo 1"
                                nValorR1 = nValorR1 + 1
                            End If
                        'FIN Agrego JOEP 20170420
                        
                        'Inicio Comentado por JOEP 20170420
                            'If Trim(prsCredGarant!cClasGarant) = "GARANTIAS NO PREFERIDAS" Then
                                'psRiesgo = "Riesgo 1"
                            'End If
                        'Fin Comentado por JOEP 20170420
                        
                        oDoc.WTextBox nPosicion + a, 57, 12 + nAltoAdic, 36, prsCredGarant!cNumGarant, "F1", nFTabla, hCenter, , , 1, , , 2
                        oDoc.WTextBox nPosicion + a, 93, 12 + nAltoAdic, 60, prsCredGarant!cTpoGarant, "F1", nFTabla, hjustify, , , 1, , , 2
                        
                        oDoc.WTextBox nPosicion + a, 153, 12 + nAltoAdic, 15, IIf(prsCredGarant!nTratLegal = 1, 2, 1), "F1", nFTabla, hCenter, , , 1, , , 2 'Agrego JOEP 20170420
                        
                        'oDoc.WTextBox nPosicion + A, 153, 12 + nAltoAdic, 15, IIf(Trim(prsCredGarant!cClasGarant) = "GARANTIAS NO PREFERIDAS", 1, 2), "F1", nFTabla, hCenter, , , 1, , , 2 'Inicio Comentado por JOEP 20170420
                        
                        oDoc.WTextBox nPosicion + a, 188 - 20, 12 + nAltoAdic, 76, prsCredGarant!cDocDesc & " - Nº " & prsCredGarant!cNroDoc, "F1", nFTabla, hjustify, , , 1, , , 2
                        oDoc.WTextBox nPosicion + a, 264 - 20, 12 + nAltoAdic, 76, prsCredGarant!cDireccion, "F1", nFTabla, hjustify, , , 1, , , 2
                        oDoc.WTextBox nPosicion + a, 340 - 20, 12 + nAltoAdic, 20, prsCredGarant!cMoneda, "F1", nFTabla, hCenter, , , 1, , , 2
                        oDoc.WTextBox nPosicion + a, 367 - 27, 12 + nAltoAdic, 40, Format(nValCom, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
                        oDoc.WTextBox nPosicion + a, 407 - 27, 12 + nAltoAdic, 40, Format(nValRea, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
                        oDoc.WTextBox nPosicion + a, 447 - 27, 12 + nAltoAdic, 40, Format(nValUti, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
                        oDoc.WTextBox nPosicion + a, 487 - 27, 12 + nAltoAdic, 40, Format(nValDis, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
                        oDoc.WTextBox nPosicion + a, 527 - 27, 12 + nAltoAdic, 40, Format(nValGra, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
                        'oDoc.WTextBox nPosicion + a, 367 - 27, 12 + nAltoAdic, 40, Format(prsCredGarant!nTasacion, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
                        'oDoc.WTextBox nPosicion + a, 407 - 27, 12 + nAltoAdic, 40, Format(prsCredGarant!nRealizacion, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
                        'oDoc.WTextBox nPosicion + a, 447 - 27, 12 + nAltoAdic, 40, Format(prsCredGarant!nGravado, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
                        'oDoc.WTextBox nPosicion + a, 487 - 27, 12 + nAltoAdic, 40, Format(prsCredGarant!ndisponible, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
                        'oDoc.WTextBox nPosicion + a, 527 - 27, 12 + nAltoAdic, 40, Format(prsCredGarant!nValorGravado, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
                        oDoc.WTextBox nPosicion + a, 567 - 27, 12 + nAltoAdic, 25, Mid(prsCredGarant!cRelGarant, 1, 3) & ".", "F1", nFTabla, hCenter, , , 1, , , 2

                        'RECO20140804*********************************
                        nValComer = nValComer + nValCom
                        nValReali = nValReali + nValRea
                        nValUtiliz = nValUtiliz + nValUti
                        nValDisp = nValDisp + nValDis
                        nValGrava = nValGrava + nValGra
                        'nValComer = nValComer + prsCredGarant!nTasacion
                        'nValReali = nValReali + prsCredGarant!nRealizacion
                        'nValUtiliz = nValUtiliz + prsCredGarant!nGravado
                        'nValDisp = nValDisp + prsCredGarant!ndisponible
                        'nValGrava = nValGrava + prsCredGarant!nValorGravado
                        'RECO FIN*************************************
                        prsCredGarant.MoveNext
                        a = a + 12 + nAltoAdic
                    Next
                End If
            End If
            'nPosicion = 207
            oDoc.WTextBox nPosicion + a, 57, 12, 283, "TOTALES", "F1", nFTabla, hCenter, , , 1, , , 2
            oDoc.WTextBox nPosicion + a, 340, 12, 40, Format(nValComer, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
            oDoc.WTextBox nPosicion + a, 380, 12, 40, Format(nValReali, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
            oDoc.WTextBox nPosicion + a, 420, 12, 40, Format(nValUtiliz, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
            oDoc.WTextBox nPosicion + a, 460, 12, 40, Format(nValDisp, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
            oDoc.WTextBox nPosicion + a, 500, 12, 40, Format(nValGrava, gcFormView), "F1", nFTabla, hRight, , , 1, , , 2
            oDoc.WTextBox nPosicion + a, 540, 12, 25, "", "F1", nFTabla, hLeft, , , 1, , , 2
            nPosicion = nPosicion + a + 20
            i = 0
            a = 0
    'FIN SECCION Nº 3
    'SECCION Nº 4
    'FIN SECCION Nº 4
    'SECCION Nº 5
        'RECO20140619************************
            Dim nPosAmp As Integer
            oDoc.WTextBox nPosicion + 2, 55, 60, 490, "COBERTURA DE GARANTIA", "F4", lnFontSizeBody, hLeft
            nPosAmp = nPosicion
            nPosicion = nPosicion + 12
            
            oDoc.WTextBox nPosicion, 55, 12, 95, "Cobertura Exp. Este Crédito", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox nPosicion, 150, 12, 95, "Cobertura Exp. Riesgo Único", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox nPosicion, 245, 12, 95, "Tipo de Riesgo", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            nPosicion = nPosicion + 12
            If lnMontoExpEstCred = 0 Then
                oDoc.WTextBox nPosicion + a, 55, 12, 95, Format(0, gcFormView), "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            Else
                oDoc.WTextBox nPosicion + a, 55, 12, 95, Format(nValDisp / lnMontoExpEstCred, gcFormView), "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            End If
            If lnMontoRiesgoUnico = 0 Then
                oDoc.WTextBox nPosicion + a, 150, 12, 95, Format(0, gcFormView), "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            Else
                oDoc.WTextBox nPosicion + a, 150, 12, 95, Format(nValDisp / lnMontoRiesgoUnico, gcFormView), "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            End If
            
            'INICION JOEP 20170420
            If (nValorR1 >= 1) Then
                nRunico = 1
                oDoc.WTextBox nPosicion + a, 245, 12, 95, "RIESGO 1", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            Else
                oDoc.WTextBox nPosicion + a, 245, 12, 95, UCase(psRiesgo), "F1", nFTablaCabecera, hCenter, , , 1, , , 2
                'CTI3 (ferimoro)
                If psRiesgo = "Riesgo 1" Or psRiesgo = "RIESGO 1" Then
                    nRunico = 1
                Else
                  If psRiesgo = "Riesgo 2" Or psRiesgo = "RIESGO 2" Then
                    nRunico = 2
                  End If
                End If
            End If
            
            'oDoc.WTextBox nPosicion + A, 245, 12, 95, UCase(psRiesgo), "F1", nFTablaCabecera, hCenter, , , 1, , , 2 'Comentado por JOEP 20170420
            
            nPosicion = nPosicion + 25
        'END RECO***************************
        
        If Not (prsCredAmp.BOF And prsCredAmp.EOF) Then
            oDoc.WTextBox nPosAmp + 2, 360, 60, 490, "AMPLIACIÓN DE CRÉDITO", "F4", lnFontSizeBody, hLeft
            
            
            nPosAmp = nPosAmp + 12
            oDoc.WTextBox nPosAmp, 360, 12, 77, "Crédito Nº", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox nPosAmp, 437, 12, 60, "Saldo Capital", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            nPosAmp = nPosAmp + 12
            For i = 1 To prsCredAmp.RecordCount
                oDoc.WTextBox nPosAmp + a, 360, 12, 77, prsCredAmp!cCtaCodAmp, "F1", nFTablaCabecera, hCenter, , , 1, , , 2
                oDoc.WTextBox nPosAmp + a, 437, 12, 60, Format(prsCredAmp!nMonto, gcFormView), "F1", nFTablaCabecera, hRight, , , 1, , , 2
                prsCredAmp.MoveNext
                a = a + 12
            Next
            nPosicion = nPosAmp + a + 10
        End If
    'FIN SECCION Nº 5
    'SECCION Nº 6
        Dim nCanIFs As Integer
        Dim nAdicionaFila As Integer
        If Not (pRRelaBcos.EOF And pRRelaBcos.BOF) Then
            nCanIFs = pRRelaBcos.RecordCount
        End If
        If nCanIFs > 5 Then
            nAdicionaFila = nAdicionaFila + 15
        End If
        If nCanIFs > 7 Then
            nAdicionaFila = nAdicionaFila + 15
        End If
        If nCanIFs > 9 Then
            nAdicionaFila = nAdicionaFila + 15
        End If
        nPosicion = nPosicion
        oDoc.WTextBox nPosicion + 6, 55, 60, 490, "RATIOS FINANCIEROS", "F4", lnFontSizeBody, hLeft
        oDoc.WTextBox nPosicion + 16, 55, 56 + nAdicionaFila, 240, "", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 16, 295, 56 + nAdicionaFila, 240, "", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 16, 295, 12, 139, "Institución", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 16, 434, 12, 28, "Moneda", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 16, 462, 12, 40, "Saldo", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 16, 502, 12, 33, "Relacion", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        nPosicion = nPosicion + 9
        oDoc.WTextBox nPosicion + 12, 57, 12, 63, "Liquidez", "F1", 5, hLeft, , , , , , 2
        'ENDEUDADAMIENTO CON OTRAS IFIS**************************************************
           Dim nIndx As Integer
            Dim nTmpPosic As Integer
            nTmpPosic = nPosicion + 8
            
             For nIndx = 1 To pRRelaBcos.RecordCount
                'If Len(pRRelaBcos!Nombre) > 30 Then 'COMENTADO POR APRI 20170719 - MEJORA
                    oDoc.WTextBox nTmpPosic + 9, 295, 9, 139, pRRelaBcos!Nombre, "F1", 5, hLeft, , , , , , 2
                'Else 'COMENTADO POR APRI 20170719 - MEJORA
                    'oDoc.WTextBox nTmpPosic + 9, 295, 9, 139, pRRelaBcos!Nombre, "F1", nFTablaCabecera, hLeft, , , , , , 2 'COMENTADO POR APRI 20170719 - MEJORA
                'End If 'COMENTADO POR APRI 20170719 - MEJORA
                oDoc.WTextBox nTmpPosic + 9, 434, 9, 28, pRRelaBcos!Moneda, "F1", 5, hCenter, , , , , , 2
                oDoc.WTextBox nTmpPosic + 9, 462, 9, 40, Format(pRRelaBcos!Saldo, gcFormView), "F1", 5, hRight, , , , , , 2
                oDoc.WTextBox nTmpPosic + 9, 502, 9, 33, Mid(pRRelaBcos!Relacion, 1, 3) & ".", "F1", 5, hLeft, , , , , , 2
                'pRRelaBcos.MoveNext
                'APRI 20170719 - MEJORA
                If Len(pRRelaBcos!Nombre) > 42 Then
                nTmpPosic = nTmpPosic + 12
                Else
                 nTmpPosic = nTmpPosic + 8
                End If
                pRRelaBcos.MoveNext
                'END APRI
            Next
        'End If
        '********************************************************************************
'        oDoc.WTextBox nPosicion + 12, 120, 12, 63, "N/A", "F1", nFTablaCabecera, hLeft, , , 1, , , 2
'        oDoc.WTextBox nPosicion + 12, 188, 12, 63, "Patrimonio", "F1", 5, hLeft, , , , , , 2
'        oDoc.WTextBox nPosicion + 12, 221, 12, 63, "N/A", "F1", nFTablaCabecera, hLeft, , , 1, , , 2
'        nPosicion = nPosicion + 12
'        oDoc.WTextBox nPosicion + 12, 57, 12, 63, "Endeudamiento Patrimonial", "F1", 5, hLeft, , , , , , 2
'        oDoc.WTextBox nPosicion + 12, 120, 12, 63, "N/A", "F1", nFTablaCabecera, hLeft, , , 1, , , 2
'        oDoc.WTextBox nPosicion + 12, 188, 12, 63, "Inventario", "F1", 5, hLeft, , , , , , 2
'        oDoc.WTextBox nPosicion + 12, 221, 12, 63, "N/A", "F1", nFTablaCabecera, hLeft, , , 1, , , 2
'        nPosicion = nPosicion + 12
'        oDoc.WTextBox nPosicion + 12, 57, 12, 63, "Rentabilidad Patrimonial", "F1", 5, hLeft, , , , , , 2
'        oDoc.WTextBox nPosicion + 12, 120, 12, 63, "N/A", "F1", nFTablaCabecera, hLeft, , , 1, , , 2
'        oDoc.WTextBox nPosicion + 12, 188, 12, 63, "Excedente", "F1", 5, hLeft, , , , , , 2
'        oDoc.WTextBox nPosicion + 12, 221, 12, 63, Format(lnExcedente, "0.00"), "F1", nFTablaCabecera, hLeft, , , 1, , , 2
'        nPosicion = nPosicion + 12
'        oDoc.WTextBox nPosicion + 12, 57, 12, 63, "Capacid. De Pago", "F1", 5, hLeft, , , , , , 2
'        oDoc.WTextBox nPosicion + 12, 120, 12, 63, Format(lnCapacidadPago * 100, "0.00") & "%", "F1", nFTablaCabecera, hRight, , , 1, , , 2
        'DATOS RATIOS FINANCIEROS FORMATOS EVALUACION *********************************************************************
        oDoc.WTextBox nPosicion + 12, 120, 12, 63, "N/A", "F1", nFTablaCabecera, hLeft, , , 1, , , 2
        oDoc.WTextBox nPosicion + 12, 188, 12, 63, "Patrimonio", "F1", 5, hLeft, , , , , , 2
        oDoc.WTextBox nPosicion + 12, 221, 12, 63, "N/A", "F1", nFTablaCabecera, hLeft, , , 1, , , 2
        nPosicion = nPosicion + 12
        oDoc.WTextBox nPosicion + 12, 57, 12, 63, "Endeudamiento Patrimonial", "F1", 5, hLeft, , , , , , 2
        oDoc.WTextBox nPosicion + 12, 120, 12, 63, "N/A", "F1", nFTablaCabecera, hLeft, , , 1, , , 2
        oDoc.WTextBox nPosicion + 12, 188, 12, 63, "Inventario", "F1", 5, hLeft, , , , , , 2
        oDoc.WTextBox nPosicion + 12, 221, 12, 63, "N/A", "F1", nFTablaCabecera, hLeft, , , 1, , , 2
        nPosicion = nPosicion + 12
        oDoc.WTextBox nPosicion + 12, 57, 12, 63, "Rentabilidad Patrimonial", "F1", 5, hLeft, , , , , , 2
        oDoc.WTextBox nPosicion + 12, 120, 12, 63, "N/A", "F1", nFTablaCabecera, hLeft, , , 1, , , 2
        oDoc.WTextBox nPosicion + 12, 188, 12, 63, "Excedente", "F1", 5, hLeft, , , , , , 2
        oDoc.WTextBox nPosicion + 12, 221, 12, 63, IIf(pDRRatiosF!nExceMensual = 0, "N/A", Format(pDRRatiosF!nExceMensual, "0.00")), "F1", nFTablaCabecera, hLeft, , , 1, , , 2
        nPosicion = nPosicion + 12
        oDoc.WTextBox nPosicion + 12, 57, 12, 63, "Capacid. De Pago", "F1", 5, hLeft, , , , , , 2
        oDoc.WTextBox nPosicion + 12, 120, 12, 63, IIf(pDRRatiosF!nCapPagNeta = 0, "N/A", Format(pDRRatiosF!nCapPagNeta * 100, "0.00") & "%"), "F1", nFTablaCabecera, hRight, , , 1, , , 2
        'FIN DATOS RATIOS FORMATOS EVALUACION **********************************************************************************
        oDoc.WTextBox nPosicion + 12, 188, 12, 63, "Cuota", "F1", 5, hLeft, , , , , , 2
        oDoc.WTextBox nPosicion + 12, 221, 12, 63, Format(lnCuota, "0.00"), "F1", nFTablaCabecera, hLeft, , , 1, , , 2
        nPosicion = nPosicion + 12 + 20 + nAdicionaFila
    'FIN SECCION Nº 6
    'SECCION Nº 7
        oDoc.WTextBox nPosicion, 55, 60, 490, "CALIFICACIÓN Y RELACIÓN DE TITULARES / CÓNYUGE / AVALES", "F4", 7, hLeft
        oDoc.WTextBox nPosicion + 12, 55, 56, 480, "", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 12, 55, 12, 240, "Nombre", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 12, 295, 12, 115, "Relación", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 12, 410, 12, 25, "Normal", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 12, 435, 12, 25, "Poten.", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 12, 460, 12, 25, "Defic.", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 12, 485, 12, 25, "Dudos", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 12, 510, 12, 25, "Pérdida", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        'CALIFICACION Y RELACION DE TITULARES/CONYUGUE /AVALES**************************************************
        If Not (prsCalfSBSRela.EOF And prsCalfSBSRela.BOF) Then
            Dim nIndx2 As Integer
            Dim nTmpPosic2 As Integer
            nTmpPosic2 = nPosicion + 14
            For nIndx2 = 1 To prsCalfSBSRela.RecordCount
                oDoc.WTextBox nTmpPosic2 + 9, 55, 9, 240, prsCalfSBSRela!cPersNombre, "F1", nFTablaCabecera, hLeft, , , , , , 2
                oDoc.WTextBox nTmpPosic2 + 9, 295, 9, 115, prsCalfSBSRela!cConsDescripcion, "F1", nFTablaCabecera, hLeft, , , , , , 2
                '*** FRHU 20160823
                'oDoc.WTextBox nTmpPosic2 + 9, 410, 9, 25, prsCalfSBSRela!Normal & "%", "F1", nFTablaCabecera, hCenter, , , , , , 2
                'oDoc.WTextBox nTmpPosic2 + 9, 435, 9, 25, prsCalfSBSRela!Potencial & "%", "F1", nFTablaCabecera, hCenter, , , , , , 2
                'oDoc.WTextBox nTmpPosic2 + 9, 460, 9, 25, prsCalfSBSRela!DEFICIENTE & "%", "F1", nFTablaCabecera, hCenter, , , , , , 2
                'oDoc.WTextBox nTmpPosic2 + 9, 485, 9, 25, prsCalfSBSRela!DUDOSO & "%", "F1", nFTablaCabecera, hCenter, , , , , , 2
                'oDoc.WTextBox nTmpPosic2 + 9, 510, 9, 25, prsCalfSBSRela!PERDIDA & "%", "F1", nFTablaCabecera, hCenter, , , , , , 2
                oDoc.WTextBox nTmpPosic2 + 9, 410, 9, 28, prsCalfSBSRela!Normal & "%", "F1", nFTablaCabecera, hCenter, , , , , , 2
                oDoc.WTextBox nTmpPosic2 + 9, 435, 9, 28, prsCalfSBSRela!Potencial & "%", "F1", nFTablaCabecera, hCenter, , , , , , 2
                oDoc.WTextBox nTmpPosic2 + 9, 460, 9, 28, prsCalfSBSRela!DEFICIENTE & "%", "F1", nFTablaCabecera, hCenter, , , , , , 2
                oDoc.WTextBox nTmpPosic2 + 9, 485, 9, 28, prsCalfSBSRela!DUDOSO & "%", "F1", nFTablaCabecera, hCenter, , , , , , 2
                oDoc.WTextBox nTmpPosic2 + 9, 510, 9, 28, prsCalfSBSRela!PERDIDA & "%", "F1", nFTablaCabecera, hCenter, , , , , , 2
                '*** FIN FRHU 20160823
                prsCalfSBSRela.MoveNext
                nTmpPosic2 = nTmpPosic2 + 8
            Next
        End If
    'FIN SECCION Nº 7
    'APRI20170719 TI-ERS025-2017
    If prsPersVinc.RecordCount > 0 Then
        nPosicion = nPosicion + 70
        oDoc.WTextBox nPosicion, 55, 60, 490, "CALIFICACIÓN DE VINCULADOS AL TRABAJADOR", "F4", 7, hLeft
        oDoc.WTextBox nPosicion + 9, 55, 56, 480, "", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 9, 55, 12, 240, "Nombre", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 9, 295, 12, 115, "Relación", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 9, 410, 12, 25, "Normal", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 9, 435, 12, 25, "Poten.", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 9, 460, 12, 25, "Defic.", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 9, 485, 12, 25, "Dudos", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion + 9, 510, 12, 25, "Pérdida", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        If Not (prsPersVinc.EOF And prsPersVinc.BOF) Then
            Dim nIndxP As Integer
            Dim nTmpPosicP As Integer
            nTmpPosicP = nPosicion + 14
            For nIndxP = 1 To prsPersVinc.RecordCount
                oDoc.WTextBox nTmpPosicP + 9, 55, 11, 240, prsPersVinc!cPersNombre, "F1", nFTablaCabecera, hLeft, , , , , , 1
                oDoc.WTextBox nTmpPosicP + 9, 295, 11, 115, prsPersVinc!cRelacion, "F1", nFTablaCabecera, hLeft, , , , , , 1
                oDoc.WTextBox nTmpPosicP + 9, 410, 11, 28, prsPersVinc!Normal & "%", "F1", nFTablaCabecera, hCenter, , , , , , 1
                oDoc.WTextBox nTmpPosicP + 9, 435, 11, 28, prsPersVinc!Potencial & "%", "F1", nFTablaCabecera, hCenter, , , , , , 1
                oDoc.WTextBox nTmpPosicP + 9, 460, 11, 28, prsPersVinc!DEFICIENTE & "%", "F1", nFTablaCabecera, hCenter, , , , , , 1
                oDoc.WTextBox nTmpPosicP + 9, 485, 11, 28, prsPersVinc!DUDOSO & "%", "F1", nFTablaCabecera, hCenter, , , , , , 1
                oDoc.WTextBox nTmpPosicP + 9, 510, 11, 28, prsPersVinc!PERDIDA & "%", "F1", nFTablaCabecera, hCenter, , , , , , 1
                prsPersVinc.MoveNext
                nTmpPosicP = nTmpPosicP + 8
            Next
        End If
    End If
    'END APRI
    
    'WIOR 20160621 *** SOBREENDEUDAMIENTO DE CLIENTES
    nPosicion = nPosicion + 70
    If Not (prsSobreEnd.EOF And prsSobreEnd.BOF) Then
        oDoc.WTextBox nPosicion, 55, 56, 330, "SOBREENDEUDAMIENTO DE CLIENTES", "F4", 7, hLeft
        nPosicion = nPosicion + 12
        oDoc.WTextBox nPosicion, 55, 12, 80, "Deuda Potencial", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion, 135, 12, 80, Format(CCur(prsSobreEnd!nDeudaTotal), "###," & String(15, "#") & "#0.00"), "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        If Not (prsSobreEndCodigos.EOF And prsSobreEndCodigos.BOF) Then
            nPosicion = nPosicion + 18

            oDoc.WTextBox nPosicion, 55, 12, 50, "Códigos", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox nPosicion, 105, 12, 110, "Resultados", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox nPosicion, 215, 12, 160, "Detalle", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            oDoc.WTextBox nPosicion, 375, 12, 160, "Plan de Mitigación del Riesgo", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
            nPosicion = nPosicion + 12

            For i = 1 To prsSobreEndCodigos.RecordCount
                '*** FRHU 20160823
                'oDoc.WTextBox nPosicion, 55, 15, 50, Trim(prsSobreEndCodigos!cCodigo), "F1", nFTablaCabecera, hCenter, vMiddle, , 1, , , 2
                'oDoc.WTextBox nPosicion, 105, 15, 110, Trim(prsSobreEndCodigos!cResultado), "F1", nFTablaCabecera, hLeft, vMiddle, , 1, , , 2
                'oDoc.WTextBox nPosicion, 215, 15, 160, Trim(prsSobreEndCodigos!cDetalle), "F1", 6, hLeft, vMiddle, , 1, , , 2
                'oDoc.WTextBox nPosicion, 375, 15, 160, Trim(prsSobreEndCodigos!cPlanMitigacion), "F1", 6, hLeft, vMiddle, , 1, , , 2
                'prsSobreEndCodigos.MoveNext
                'nPosicion = nPosicion + 15
                oDoc.WTextBox nPosicion, 55, 22, 50, Trim(prsSobreEndCodigos!cCodigo), "F1", nFTablaCabecera, hCenter, vMiddle, , 1, , , 2
                oDoc.WTextBox nPosicion, 105, 22, 110, Trim(prsSobreEndCodigos!cResultado), "F1", nFTablaCabecera, hLeft, vMiddle, , 1, , , 2
                oDoc.WTextBox nPosicion, 215, 22, 160, Trim(prsSobreEndCodigos!cDetalle), "F1", 6, hLeft, , , 1, , , 2
                oDoc.WTextBox nPosicion, 375, 22, 160, Trim(prsSobreEndCodigos!cPlanmitigacion), "F1", 6, hLeft, vMiddle, , 1, , , 2
                prsSobreEndCodigos.MoveNext
                nPosicion = nPosicion + 22
                'FIN FRHU 20160823
            Next i
        End If
        'nPosicion = nPosicion + 12
        nPosicion = nPosicion + 3 'FRHU 20160823
    End If
    'WIOR FIN ********
    
    'RECO NUEVO
    'nPosicion = nPosicion + 70
    If Not (prsComentAnalis.BOF And prsComentAnalis.EOF) Then
        oDoc.WTextBox nPosicion, 55, 56, 330, "COMENTARIO ANALISTA", "F4", nFTablaCabecera, hLeft
        nPosicion = nPosicion + 9
        oDoc.WTextBox nPosicion, 55, 40, 480, prsComentAnalis!cComentAnalista, "F1", nFTablaCabecera, hjustify, , , 1, , , 2
        nPosicion = nPosicion + 48
    End If
    'RECO FIN NUEVO
    'SECCION N° 9 CUADRO DE INGRESOS Y GATOS MENSUALES
            '*** FRHU 20160823: Quitado segun lo indicado por el usuario RUSI
            'oDoc.WTextBox nPosicion, 55, 12, 300, "CUADRO DE INGRESOS Y GASTOS MENSUALES", "F4", nFTablaCabecera, hLeft
            'nPosicion = nPosicion + 12
            'oDoc.WTextBox nPosicion, 55, 12, 80, "INGRESO BRUTO", "F1", nFTablaCabecera, hLeft, , , 1, , , 2
            'oDoc.WTextBox nPosicion, 135, 12, 80, Format(0, gcFormView), "F1", nFTablaCabecera, hRight, , , 1, , , 2
            'oDoc.WTextBox nPosicion, 215, 12, 80, "OTROS INGRESOS", "F1", nFTablaCabecera, hLeft, , , 1, , , 2
            'oDoc.WTextBox nPosicion, 295, 12, 80, Format(0, gcFormView), "F1", nFTablaCabecera, hRight, , , 1, , , 2
            'nPosicion = nPosicion + 12
            'oDoc.WTextBox nPosicion, 55, 12, 80, "INGRESO NETO", "F1", nFTablaCabecera, hLeft, , , 1, , , 2
            'oDoc.WTextBox nPosicion, 135, 12, 80, Format(nIngNeto, gcFormView), "F1", nFTablaCabecera, hRight, , , 1, , , 2
            'oDoc.WTextBox nPosicion, 215, 12, 80, "GASTOS FAMILIARES", "F1", nFTablaCabecera, hLeft, , , 1, , , 2
            'oDoc.WTextBox nPosicion, 295, 12, 80, Format(nGasFamiliar, gcFormView), "F1", nFTablaCabecera, hRight, , , 1, , , 2
            '*** FIN FRHU 20160823
    'nUnico , cCodHojEval
    
    'FIN SECCION N° 9
    'SECCION Nº 8
        'nPosicion = nPosicion + 12 + 10 'FRHU 20160823
        
        'oDoc.WTextBox nPosicion, 55, 56, 280, "", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion, 55, 100, 210, "", "F1", nFTablaCabecera, hCenter, , , 1, , , 2 'FRHU 20160811 Anexo002
        'oDoc.WTextBox nPosicion, 55, 56, 165, "", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        'oDoc.WTextBox nPosicion, 55, 12, 165, "EXONERACIONES", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        'oDoc.WTextBox nPosicion, 220, 12, 165, "AUTORIZACIONES", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        'oDoc.WTextBox nPosicion, 55, 12, 280, "AUTORIZACIONES", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion, 55, 12, 210, "Autorizaciones", "F1", nFTablaCabecera, hCenter, , , 1, , , 2 'FRHU 20160811 Anexo002
        'oDoc.WTextBox nPosicion, 385, 12, 200, "NIVELES DE APROBACION POR EXPOSICION", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion, 265, 12, 270, "Nivel de Aprobación", "F1", nFTablaCabecera, hCenter, , , 1, , , 2 'FRHU 20160811 Anexo002 ERS002-2016
        'oDoc.WTextBox nPosicion, 385, 56, 200, "", "F1", nFTablaCabecera, hCenter, , , 1, , , 2
        oDoc.WTextBox nPosicion, 265, 100, 270, "", "F1", nFTablaCabecera, hCenter, , , 1, , , 2 'FRHU 20160811 Anexo002
        
        nPosicion = nPosicion + 12 'FRHU 20160811 Anexo002 ERS002-2016 NEW
        oDoc.WTextBox nPosicion, 265, 12, 135, "Por Autorización", "F1", nFTablaCabecera, hCenter, , , 1, , , 2 'FRHU 20160811 Anexo002 ERS002-2016 NEW
        oDoc.WTextBox nPosicion, 400, 12, 135, "Por Exposición", "F1", nFTablaCabecera, hCenter, , , 1, , , 2 'FRHU 20160811 Anexo002 ERS002-2016 NEW
        oDoc.WTextBox nPosicion, 400, 88, 135, "", "F1", nFTablaCabecera, hCenter, , , 1, , , 2 'FRHU 20160811 Anexo002
        
        If psHabilitaNiveles = 0 Then
            If Not (prsExoAutCred.EOF And prsExoAutCred.BOF) Then
                Dim nIndx3 As Integer
                Dim nTmpPosic3 As Integer
                Dim nTmpPosicExo As Integer
                Dim nTmpPosicAut As Integer
                nTmpPosic3 = nPosicion + 14
                nTmpPosicExo = nPosicion + 2
                nTmpPosicAut = nPosicion + 2
                'For nIndx3 = 1 To 3
                For nIndx3 = 1 To prsExoAutCred.RecordCount
                 Dim texto As String
                    If prsExoAutCred!nTipoExoneraCod = 1 Then
                        Set lrsRespNvlExo = oDCredExoAut.RecuperaRespExoAut(prsExoAutCred!cExoneraCod, IIf(lnMontoRiesgoUnico = 0, lnMontoExpEstCred, lnMontoRiesgoUnico))
                        oDoc.WTextBox nTmpPosicExo + 9, 55, 9, 150, prsExoAutCred!cExoneraDesc, "F1", nFTablaCabecera, hLeft, , , , , , 2
                        
                        If Not (lrsRespNvlExo.BOF And lrsRespNvlExo.EOF) Then
                            oDoc.WTextBox nTmpPosicExo + 9, 145, 9, 150, lrsRespNvlExo!cNivAprDesc, "F1", nFTablaCabecera, hLeft, , , , , , 2
                        End If
                        nTmpPosicExo = nTmpPosicExo + 5
                    Else
                        Set lrsRespNvlExo = oDCredExoAut.RecuperaRespExoAut(prsExoAutCred!cExoneraCod, IIf(lnMontoRiesgoUnico = 0, lnMontoExpEstCred, lnMontoRiesgoUnico))
                        oDoc.WTextBox nTmpPosicAut + 9, 225, 9, 240, prsExoAutCred!cExoneraDesc, "F1", nFTablaCabecera, hLeft, , , , , , 2
                        If Not (lrsRespNvlExo.BOF And lrsRespNvlExo.EOF) Then
                            oDoc.WTextBox nTmpPosicAut + 9, 295, 9, 150, lrsRespNvlExo!cNivAprDesc, "F1", nFTablaCabecera, hLeft, , , , , , 2
                        End If
                        nTmpPosicAut = nTmpPosicAut + 5
                    End If
                    prsExoAutCred.MoveNext
                    Set lrsRespNvlExo = Nothing
                Next
            End If
            
'             If Not (prsCredResulNivApr.EOF And prsCredResulNivApr.BOF) Then
'                Dim nIndx4 As Integer
'                Dim nTmpPosic4 As Integer
'
'                nTmpPosic4 = nPosicion + 14
'                For nIndx = 1 To prsCredResulNivApr.RecordCount
'                    oDoc.WTextBox nTmpPosic4 + 12, 295, 12, 80, prsCredResulNivApr!cNivAprDesc, "F1", nFTablaCabecera, hLeft
'                    nTmpPosic4 = nTmpPosic4 + 5
'                    prsCredResulNivApr.MoveNext
'                Next
'            End If
'**************************ferimoro**************************
If pR!Modalidad = "REFINANCIADO" Then
 Dim sNivelinicial As String
 Dim sPdirecto
 Dim nTmpPosicNiv4 As Integer
     
    ' Dim oRecCtaRef As COMDCredito.DCOMCredito
    ' Dim cCtaRefe As ADODB.Recordset
 
    Set oRecCtaRef = New COMDCredito.DCOMCredito
    Set cCtaRefe = oRecCtaRef.RecuperaCctaReferencia(pR!Nro_Credito)
    Set oRecCtaRef = Nothing

 
    sNivelinicial = NivelAprobacionInicial(pR!Nro_Credito, cCtaRefe!cCtaCodRef)


    sPdirecto = pR!Tipo_Prod

    Set oDpExposicion = New COMDCredito.DCOMCredito
    Set expo = oDpExposicion.RecuperaAutoporExpo(nRunico, pR!cDestino, pR!Modalidad, lnMontoRiesgoUnico, pR!Oficina, sPdirecto, nClePref, sNivelinicial)
    Set oDpExposicion = Nothing
    
            nTmpPosicNiv4 = nPosicion + 1
            oDoc.WTextBox nTmpPosicNiv4 + 12, 402, 12, 305, IIf(IsNull(expo!autExp) = True, "", expo!autExp), "F1", nFTablaCabecera, hLeft
    
Else
                     '
    
    If pR!Tipo_Prod = "Agropecuario Directo" Then
    sPdirecto = "AGROPECUARIOS DIRECTO"
    Else
    sPdirecto = pR!Tipo_Prod
    End If
                                    
    Set oDpExposicion = New COMDCredito.DCOMCredito
    Set expo = oDpExposicion.RecuperaAutoporExpo(nRunico, pR!cDestino, pR!Modalidad, lnMontoRiesgoUnico, pR!Oficina, sPdirecto, nClePref)
    Set oDpExposicion = Nothing


            
            nTmpPosicNiv4 = nPosicion + 1
            oDoc.WTextBox nTmpPosicNiv4 + 12, 402, 12, 305, IIf(IsNull(expo!autExp) = True, "", expo!autExp), "F1", nFTablaCabecera, hLeft
End If
'************************************************************
            
            
        End If
        'FRHU 20160811 Anexo002 ERS002-2016
        If Not (prsAutorizaciones.EOF And prsAutorizaciones.BOF) Then
            Dim nTmpPosicExo5 As Integer
            nTmpPosicExo5 = nPosicion + 2
            For nIndx = 1 To prsAutorizaciones.RecordCount
                If Len(prsAutorizaciones!cExoneraDesc) <= 69 Then
                    oDoc.WTextBox nTmpPosicExo5 + 9, 55, 9, 280, prsAutorizaciones!cExoneraDesc, "F1", nFTablaCabecera, hLeft, , , , , , 2
                    oDoc.WTextBox nTmpPosicExo5 + 11, 267, 12, 305, prsAutorizaciones!cNivAprDesc, "F1", nFTablaCabecera, hLeft
                    nTmpPosicExo5 = nTmpPosicExo5 + 9
                Else
                    oDoc.WTextBox nTmpPosicExo5 + 9, 55, 9, 280, prsAutorizaciones!cExoneraDesc, "F1", nFTablaCabecera, hLeft, , , , , , 2
                    oDoc.WTextBox nTmpPosicExo5 + 11, 267, 12, 305, prsAutorizaciones!cNivAprDesc, "F1", nFTablaCabecera, hLeft
                    nTmpPosicExo5 = nTmpPosicExo5 + 18
                End If
                prsAutorizaciones.MoveNext
            Next nIndx
        End If
        'FIN FRHU 20160811
    'FIN SECCION Nº 8
    'SECCION Nº9
    'nPosicion = nPosicion + 50
    nPosicion = nPosicion + 80 'FRHU 20160811 Anexo002 ERS002-2016
    oDoc.WTextBox nPosicion + 15, 150, 56, 150, "RESOLUCION DE COMITÉ, EN CONCLUSION: ", "F1", nFTablaCabecera, Left
    nPosicion = nPosicion + 12
    oDoc.WTextBox nPosicion + 15, 150, 56, 70, "MONTO", "F1", nFTablaCabecera, Left
    oDoc.WTextBox nPosicion + 15, 250, 56, 70, "CUOTAS", "F1", nFTablaCabecera, Left
    oDoc.WTextBox nPosicion + 15, 350, 56, 70, "TI", "F1", nFTablaCabecera, Left
    oDoc.WTextBox nPosicion + 15, 450, 56, 70, "VCTO", "F1", nFTablaCabecera, Left
    
    nPosicion = nPosicion + 20
    oDoc.WTextBox nPosicion + 15, 70, 56, 150, "APROBADO POR: ", "F1", nFTablaCabecera, Left
    oDoc.WTextBox nPosicion + 15, 150, 56, 70, "...................", "F1", nFTablaCabecera, Left
    oDoc.WTextBox nPosicion + 15, 250, 56, 70, "...................", "F1", nFTablaCabecera, Left
    oDoc.WTextBox nPosicion + 15, 350, 56, 70, "...................", "F1", nFTablaCabecera, Left
    oDoc.WTextBox nPosicion + 15, 450, 56, 70, "...................", "F1", nFTablaCabecera, Left
    'FIN SECCION N9
    oDoc.PDFClose
    oDoc.Show
    Exit Sub
ErrorImprimirPDF:
    MsgBox Err.Description, vbInformation, "Aviso"
End Sub
Public Function ValorMayor(ByVal nV1 As Integer, ByVal nV2 As Integer, ByVal nV3 As Integer, ByVal nV4 As Integer) As Integer
    Dim nValM As Integer
    Dim nArreglo(4) As Integer
    Dim i As Integer, j As Integer
    
    nArreglo(0) = nV1
    nArreglo(1) = nV2
    nArreglo(2) = nV3
    nArreglo(3) = nV4
    ValorMayor = 0
    For i = 0 To 3
        Dim nContador As Integer
        For j = 0 To 3
            If nArreglo(i) >= nArreglo(j) Then
            nContador = nContador + 1
            End If
        Next
        If nContador = 4 Then
            ValorMayor = nArreglo(i)
        End If
        nContador = 0
    Next
End Function
'RECO FIN*****************************************************

'CTI3 (ferimoro)*************************************
Public Function NivelAprobacionInicial(ByVal psCtaCodAct As String, ByVal psCtaCodAnt As String) As String

    Dim oDCred As COMDCredito.DCOMCredito
    Dim oDCredExoAut As COMDCredito.DCOMNivelAprobacion
    Dim oDpExposicion As COMDCredito.DCOMCredito
    Dim R, RCredGarant, RRiesgoUnico, expo As ADODB.Recordset

    Dim sPdirecto, sRiesgo, psRiesgo As String
    Dim nValorR2, nValorR1, i, nRunico As Integer
    Dim lnMontoRiesgoUnico As Double
    
    nValorR1 = 0
    nValorR2 = 0
    
    
    Set oDCred = New COMDCredito.DCOMCredito
    Set R = oDCred.RecuperaDatosAprobacionCreditos(psCtaCodAnt)
    Set RCredGarant = oDCred.ObtieneDatasGarantiaCred(psCtaCodAnt)
    
    
    Set oDCredExoAut = New COMDCredito.DCOMNivelAprobacion
    Set RRiesgoUnico = oDCredExoAut.ObtineRiesgoUnicoCred(psCtaCodAnt)
    
    If Not (RRiesgoUnico.EOF And RRiesgoUnico.BOF) Then
        lnMontoRiesgoUnico = RRiesgoUnico!nMonto
    End If
    
    If val(Mid(psCtaCodAnt, 6, 3)) >= 800 And val(Mid(psCtaCodAnt, 6, 3)) < 900 Then
        sRiesgo = "RIESGO 2"
    Else
        sRiesgo = "RIESGO 1"
    End If
    
    
    If Not (RCredGarant.BOF And RCredGarant.EOF) Then
        If RCredGarant.RecordCount > 0 Then
            For i = 1 To RCredGarant.RecordCount
                If RCredGarant!nTratLegal = 1 Then
                    psRiesgo = "Riesgo 2"
                    nValorR2 = nValorR2 + 1
                Else
                    psRiesgo = "Riesgo 1"
                    nValorR1 = nValorR1 + 1
                End If
             RCredGarant.MoveNext
            Next i
        End If
    End If
    
    If (nValorR1 >= 1) Then
                nRunico = 1
            Else
                If psRiesgo = "Riesgo 1" Then
                    nRunico = 1
                Else
                  If psRiesgo = "Riesgo 2" Then
                    nRunico = 2
                  End If
                End If
    End If
    
    
    
    If R!Tipo_Prod = "Agropecuario Directo" Then
        sPdirecto = "AGROPECUARIOS DIRECTO"
    Else
        sPdirecto = R!Tipo_Prod
    End If
                                    
    Set oDpExposicion = New COMDCredito.DCOMCredito
    Set expo = oDpExposicion.RecuperaAutoporExpo(nRunico, R!cDestino, R!Modalidad, lnMontoRiesgoUnico, R!Oficina, sPdirecto)
    Set oDpExposicion = Nothing
    
    NivelAprobacionInicial = expo!autExp
    
End Function



Private Sub ImprimirPdfCartillaAutorizacion() 'add PTI1 ERS070-2018 11/12/2018
    Dim sParrafoUno As String
    Dim sParrafoDos As String
    Dim sParrafoTres As String
    Dim sParrafoCuatro As String
    Dim sParrafoCinco As String
    Dim sParrafoSeis As String
    Dim sParrafoSiete As String
    Dim sParrafoOcho As String
    Dim oDoc As cPDF
    Dim nAltura As Integer
    
    Set oDoc = New cPDF
    'Creación del Archivo
    oDoc.Author = gsCodUser
    oDoc.Creator = "SICMACT - Negocio"
    oDoc.Producer = "Caja Municipal de Ahorros y Crédito de Maynas S.A."
    oDoc.Subject = "Cartilla Autorización y Actualización de datos personales"
    oDoc.Title = "Cartilla Autorización y Actualización de datos personales"
    
    If Not oDoc.PDFCreate(App.Path & "\Spooler\CartillaAutorizacionActualizacionDeDatos" & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
        Exit Sub
    End If
    
    oDoc.Fonts.Add "F1", "Arial", TrueType, Normal, WinAnsiEncoding
    oDoc.Fonts.Add "F2", "Arial", TrueType, Bold, WinAnsiEncoding
    
    'oDoc.LoadImageFromFile App.path & "\logo_cmacmaynas.bmp", "Logo"
    oDoc.LoadImageFromFile App.Path & "\Logo_2015.jpg", "Logo" 'O
    
    'Tamaño de hoja A4
    oDoc.NewPage A4_Vertical
    '<body>
    nAltura = 20
    oDoc.WTextBox 10, 10, 780, 575, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
    'oDoc.WImage 60, 480, 35, 100, "Logo"
    oDoc.WImage 70, 460, 50, 100, "Logo" 'O
    oDoc.WTextBox 90, 50, 15, 500, "AUTORIZACIÓN PARA EL TRATAMIENTO DE DATOS PERSONALES", "F2", 11, hCenter 'agregado por pti1 ers070-2018 05/12/2018
     
    oDoc.WTextBox 125, 56, 360, 520, (MatPersona(1).sNombres & " " & MatPersona(1).sApePat & " " & MatPersona(1).sApeMat & IIf(Len(MatPersona(1).sApeCas) = 0, "", " " & IIf(MatPersona(1).sEstadoCivil = "2", "DE", "VDA") & " " & MatPersona(1).sApeCas)), "F1", 11, hjustify
    oDoc.WTextBox 125, 484, 360, 520, (Trim(MatPersona(1).sPersIDnro)), "F1", 11, hjustify
    oDoc.WTextBox 125, 56, 360, 520, ("___________________________________________________________"), "F1", 11, hjustify
    oDoc.WTextBox 125, 481, 360, 520, ("____________"), "F1", 11, hjustify
    oDoc.WTextBox 125, 35, 10, 520, ("Yo, " & String(120, vbTab) & "  con DOI N° " & String(22, vbTab) & ""), "F1", 11, hjustify

    'sParrafoUno = "Yo " & String(78, vbTab) & "  con DOI N° " & String(20, vbTab) & "autorizo y otorgo por tiempo  "
      sParrafoUno = "autorizo y otorgo por tiempo indefinido, " & String(0.52, vbTab) & "mi consentimiento libre, previo, expreso, inequívoco e informado a" & Chr$(13) & _
                   "la " & String(0.52, vbTab) & "CAJA MUNICIPAL DE AHORRO Y CRÉDITO DE MAYNAS " & String(0.52, vbTab) & "S.A. " & String(0.52, vbTab) & "(en " & String(0.52, vbTab) & "adelante," & String(0.52, vbTab) & " ""LA CAJA""), " & String(0.51, vbTab) & " para " & String(0.51, vbTab) & " el" & Chr$(13) & _
                   "tratamiento de mis datos personales proporcionados " & String(0.7, vbTab) & " en contexto de la contratación de cualquier producto " & Chr$(13) & _
                   "(activo y/o pasivo)" & String(0.52, vbTab) & " o" & String(0.51, vbTab) & " servicio, " & String(0.52, vbTab) & " así " & String(0.52, vbTab) & "como " & String(0.52, vbTab) & "resultado" & String(0.52, vbTab) & "de " & String(0.52, vbTab) & " la suscripción de contratos, " & String(0.52, vbTab) & " formularios, " & String(0.52, vbTab) & " y a los " & Chr$(13) & _
                   "recopilados anteriormente, actualmente y/o por recopilar por " & String(0.52, vbTab) & "LA CAJA. " & String(0.53, vbTab) & "Asimismo, " & String(0.53, vbTab) & "otorgo " & String(0.53, vbTab) & "mi autorización" & Chr$(13) & _
                   "para el envío de información  promocional y/o publicitaria de los servicios y productos que" & String(0.53, vbTab) & " LA CAJA ofrece, " & Chr$(13) & _
                   "a tráves de cualquier medio de comunicación que se considere apropiado para su difusión, " & String(0.53, vbTab) & "y " & String(0.52, vbTab) & "para" & String(0.53, vbTab) & " su uso " & Chr$(13) & _
                   "en la gestión administrativa " & String(0.53, vbTab) & " y " & String(0.5, vbTab) & " comercial de  " & String(0.53, vbTab) & "LA  " & String(0.53, vbTab) & "CAJA " & String(0.53, vbTab) & " que guarde relación con su objeto social.  " & String(0.53, vbTab) & "En " & String(0.52, vbTab) & "ese " & Chr$(13) & _
                   "sentido, autorizo a LA CAJA al uso de mis datos personales para tratamientos que supongan el " & String(0.52, vbTab) & "desarrollo" & Chr$(13) & _
                   "de acciones y actividades comerciales, incluyendo la realización de estudios  de  mercado, " & String(0.53, vbTab) & " elaboración " & String(0.52, vbTab) & "de" & Chr$(13) & _
                   "perfiles de compra " & String(0.53, vbTab) & " y evaluaciones financieras. " & String(0.54, vbTab) & " El uso y tratamiento de mis datos personales, " & String(0.54, vbTab) & "se sujetan" & Chr$(13) & _
                   "a lo establecido por el artículo 13° de la Ley N° 29733 - Ley de Protección de Datos Personales."
    
  
    sParrafoDos = "Declaro conocer el compromiso de " & String(0.52, vbTab) & "LA CAJA " & String(0.52, vbTab) & " por garantizar el mantenimiento de la confidencialidad" & String(0.52, vbTab) & " y " & String(0.52, vbTab) & "el " & Chr$(13) & _
                  "tratamiento seguro de mis datos personales, incluyendo el resguardo en las transferencias de " & String(0.52, vbTab) & "los mismos, " & Chr$(13) & _
                  "que se realicen " & String(0.53, vbTab) & "en cumplimiento de la " & String(0.55, vbTab) & " Ley N° 29733 - Ley de Protección " & String(0.53, vbTab) & " de Datos Personales. De" & String(0.53, vbTab) & "igual " & Chr$(13) & _
                  "manera, declaro " & String(0.52, vbTab) & "conocer que los datos personales " & String(0.55, vbTab) & "proporcionados por mi persona serán incorporados " & String(0.52, vbTab) & "al " & Chr$(13) & _
                  "Banco de Datos de Clientes de  " & String(0.6, vbTab) & " LA CAJA, el cual  " & String(0.55, vbTab) & "se encuentra debidamente registrado ante la" & String(0.52, vbTab) & " Dirección " & Chr$(13) & _
                  "Nacional  " & String(0.55, vbTab) & " de  " & String(0.55, vbTab) & " Protección de Datos " & String(0.55, vbTab) & "Personales, para lo cual " & String(0.55, vbTab) & " autorizo a LA CAJA " & String(0.52, vbTab) & "que " & String(0.55, vbTab) & " recopile, registre, " & Chr$(13) & _
                  "organice, " & String(0.55, vbTab) & "almacene, " & String(0.55, vbTab) & "conserve, bloquee, suprima, extraiga, consulte, utilice, transfiera, exporte, importe" & String(0.52, vbTab) & " o " & Chr$(13) & _
                  "procese de cualquier otra forma mis datos personales, con las limitaciones que prevé la Ley."
                 
                 
    sParrafoTres = "Del mismo modo, y siempre que así lo estime necesario, declaro conocer que podré ejercitar mis derechos " & Chr$(13) & _
                   "de " & String(0.55, vbTab) & " acceso, " & String(0.56, vbTab) & " rectificación, " & String(0.58, vbTab) & " cancelación " & String(0.55, vbTab) & " y " & String(0.55, vbTab) & " oposición relativos a este tratamiento, de conformidad " & String(0.52, vbTab) & "con lo " & Chr$(13) & _
                   "establecido" & String(0.51, vbTab) & " en " & String(0.5, vbTab) & "el " & String(0.6, vbTab) & " Titulo" & String(0.54, vbTab) & " III " & String(0.54, vbTab) & " de la Ley N° 29733 - Ley de Protección de Datos " & String(0.52, vbTab) & " Personales" & String(0.52, vbTab) & " acercándome " & Chr$(13) & _
                   "a cualquiera de las Agencias de LA CAJA a nivel nacional."

   sParrafoCuatro = "Asimismo, " & String(1.4, vbTab) & " declaro " & String(1.4, vbTab) & " conocer " & String(1.4, vbTab) & " el " & String(1.4, vbTab) & "compromiso " & String(1.4, vbTab) & " de " & String(1.4, vbTab) & " LA " & String(1.4, vbTab) & "CAJA " & String(1.4, vbTab) & " por " & String(1.4, vbTab) & "respetar " & String(1.4, vbTab) & "los " & String(1.4, vbTab) & "principios " & String(1.4, vbTab) & "de " & String(1.4, vbTab) & " legalidad, " & Chr$(13) & _
                    "consentimiento, finalidad, proporcionalidad, calidad, disposición de recurso, y nivel de protección adecuado," & Chr$(13) & _
                    "conforme lo dispone la Ley N° 29733 - Ley de Protección de Datos Personales," & String(1.4, vbTab) & " para " & String(1.4, vbTab) & "el " & String(1.4, vbTab) & "tratamiento de los" & Chr$(13) & _
                    "datos personales otorgados por mi persona."
                  
    sParrafoCinco = "Esta autorización es" & String(1.5, vbTab) & " indefinida y se mantendrá inclusive" & String(0.5, vbTab) & " después de terminada(s) la(s) operación(es)" & String(0.52, vbTab) & " y/o " & Chr$(13) & _
                    "el(los) Contrato(s) que tenga" & String(1.5, vbTab) & " o pueda tener con LA CAJA" & String(1.3, vbTab) & " sin perjuicio de " & String(0.5, vbTab) & "poder ejercer mis derechos " & String(0.52, vbTab) & "de " & Chr$(13) & _
                    "acceso, rectificación, cancelación y oposición mencionados en el presente documento."
     
     Dim cfecha  As String 'pti1 add
      
     cfecha = Choose(Month(dfreg), "Enero", "Febrero", "Marzo", "Abril", _
                                        "Mayo", "Junio", "Julio", "Agosto", _
                                        "Setiembre", "Octubre", "Noviembre", "Diciembre")
            Dim nTamanio As Integer
            Dim Spac As Integer
            Dim Index As Integer
            Dim Princ As Integer
            Dim CantCarac As Integer
            Dim txtcDescrip As String
            Dim contador As Integer
            Dim nCentrar As Integer
            Dim nTamLet As Integer
            Dim spacvar As Integer
            
             nTamanio = Len(sParrafoUno)
            spacvar = 23
            Spac = 138
            Index = 1
            Princ = 1
            CantCarac = 0
            
            nTamLet = 6: contador = 0: nCentrar = 80
            
                  Do While Index <= nTamanio And Spac <> 240
                    If InStr(Index, sParrafoUno, Chr$(13), vbTextCompare) <> 0 Then
                        CantCarac = InStr(Index, sParrafoUno, Chr$(13), vbTextCompare)
                        CantCarac = CantCarac - Index
                        txtcDescrip = Mid(sParrafoUno, Index, CantCarac)
                        oDoc.WTextBox Spac, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = InStr(Index, sParrafoUno, Chr$(13), vbTextCompare) + 1
                        Spac = Spac + 5 + IIf((Len(txtcDescrip) / 50) > 1, ((Round(Len(txtcDescrip) / 50)) * 6) - 4, 0)
                        
                    ElseIf (Index <= nTamanio) And Index <> 1 Then
                        txtcDescrip = Mid(sParrafoUno, Index, nTamanio)
                        oDoc.WTextBox Spac + contador, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = nTamanio + 1
                    Else
                        oDoc.WTextBox Spac + contador, 35, 11, 520, sParrafoUno, "F1", 11
                        Index = nTamanio + 1
                    End If
            Loop
            
            nTamanio = Len(sParrafoDos)
            Spac = Spac + spacvar
            Index = 1
            Princ = 1
            CantCarac = 0
             nTamLet = 6: contador = 0: nCentrar = 80
                  Do While Index <= nTamanio And Spac <> 240
                    If InStr(Index, sParrafoDos, Chr$(13), vbTextCompare) <> 0 Then
                        CantCarac = InStr(Index, sParrafoDos, Chr$(13), vbTextCompare)
                        CantCarac = CantCarac - Index
                        txtcDescrip = Mid(sParrafoDos, Index, CantCarac)
                        oDoc.WTextBox Spac + contador, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = InStr(Index, sParrafoDos, Chr$(13), vbTextCompare) + 1
                        Spac = Spac + 5 + IIf((Len(txtcDescrip) / 50) > 1, ((Round(Len(txtcDescrip) / 50)) * 6) - 4, 0)
                        
                    ElseIf (Index <= nTamanio) And Index <> 1 Then
                        txtcDescrip = Mid(sParrafoDos, Index, nTamanio)
                        oDoc.WTextBox Spac + contador, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = nTamanio + 1
                    Else
                        oDoc.WTextBox Spac + contador, 35, 11, 520, sParrafoDos, "F1", 11
                        Index = nTamanio + 1
                    End If
            Loop
            
            nTamanio = Len(sParrafoTres)
            Spac = Spac + spacvar
            Index = 1
            Princ = 1
            CantCarac = 0
             nTamLet = 6: contador = 0: nCentrar = 80
                  Do While Index <= nTamanio And Spac <> 240
                    If InStr(Index, sParrafoTres, Chr$(13), vbTextCompare) <> 0 Then
                        CantCarac = InStr(Index, sParrafoTres, Chr$(13), vbTextCompare)
                        CantCarac = CantCarac - Index
                        txtcDescrip = Mid(sParrafoTres, Index, CantCarac)
                        oDoc.WTextBox Spac + contador, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = InStr(Index, sParrafoTres, Chr$(13), vbTextCompare) + 1
                        Spac = Spac + 5 + IIf((Len(txtcDescrip) / 50) > 1, ((Round(Len(txtcDescrip) / 50)) * 6) - 4, 0)
                        
                    ElseIf (Index <= nTamanio) And Index <> 1 Then
                        txtcDescrip = Mid(sParrafoTres, Index, nTamanio)
                        oDoc.WTextBox Spac + contador, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = nTamanio + 1
                    Else
                        oDoc.WTextBox Spac + contador, 35, 11, 520, sParrafoTres, "F1", 11
                        Index = nTamanio + 1
                    End If
            Loop
            
            nTamanio = Len(sParrafoCuatro)
            Spac = Spac + spacvar
            Index = 1
            Princ = 1
            CantCarac = 0
             nTamLet = 6: contador = 0: nCentrar = 80
                  Do While Index <= nTamanio And Spac <> 240
                    If InStr(Index, sParrafoCuatro, Chr$(13), vbTextCompare) <> 0 Then
                        CantCarac = InStr(Index, sParrafoCuatro, Chr$(13), vbTextCompare)
                        CantCarac = CantCarac - Index
                        txtcDescrip = Mid(sParrafoCuatro, Index, CantCarac)
                        oDoc.WTextBox Spac + contador, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = InStr(Index, sParrafoCuatro, Chr$(13), vbTextCompare) + 1
                        Spac = Spac + 5 + IIf((Len(txtcDescrip) / 50) > 1, ((Round(Len(txtcDescrip) / 50)) * 6) - 4, 0)
                        
                    ElseIf (Index <= nTamanio) And Index <> 1 Then
                        txtcDescrip = Mid(sParrafoCuatro, Index, nTamanio)
                        oDoc.WTextBox Spac + contador, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = nTamanio + 1
                    Else
                        oDoc.WTextBox Spac + contador, 35, 11, 520, sParrafoCuatro, "F1", 11
                        Index = nTamanio + 1
                    End If
            Loop
            
            
            nTamanio = Len(sParrafoCinco)
            Spac = Spac + spacvar
            Index = 1
            Princ = 1
            CantCarac = 0
             nTamLet = 6: contador = 0: nCentrar = 80
                  Do While Index <= nTamanio And Spac <> 240
                    If InStr(Index, sParrafoCinco, Chr$(13), vbTextCompare) <> 0 Then
                        CantCarac = InStr(Index, sParrafoCinco, Chr$(13), vbTextCompare)
                        CantCarac = CantCarac - Index
                        txtcDescrip = Mid(sParrafoCinco, Index, CantCarac)
                        oDoc.WTextBox Spac + contador, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = InStr(Index, sParrafoCinco, Chr$(13), vbTextCompare) + 1
                        Spac = Spac + 5 + IIf((Len(txtcDescrip) / 50) > 1, ((Round(Len(txtcDescrip) / 50)) * 6) - 4, 0)
                        
                    ElseIf (Index <= nTamanio) And Index <> 1 Then
                        txtcDescrip = Mid(sParrafoCinco, Index, nTamanio)
                        oDoc.WTextBox Spac + contador, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = nTamanio + 1
                    Else
                        oDoc.WTextBox Spac + contador, 35, 11, 520, sParrafoCinco, "F1", 11
                        Index = nTamanio + 1
                    End If
            Loop
            
           
    
                  'oDoc.WTextBox 136, 35, 120, 520, sParrafoUno, "F1", 11, hjustify, , , 4, vbBlue
                  'oDoc.WTextBox 310, 35, 88, 520, sParrafoDos, "F1", 11, hjustify, , , 1, vbBlue
                  'oDoc.WTextBox 430, 35, 44, 520, sParrafoTres, "F1", 11, hjustify, , , 1, vbBlue
                  'oDoc.WTextBox 485, 35, 44, 520, sParrafoCuatro, "F1", 11, hjustify, , , 1, vbBlue
                  'oDoc.WTextBox 560, 35, 33, 520, sParrafoCinco, "F1", 11, hjustify
    


    oDoc.WTextBox 610, 35, 60, 300, ("En " & sAgeReg & " a los " & Day(dfreg) & " días del mes de " & cfecha & " de " & Year(dfreg)) & ".", "F1", 11, hLeft 'O  agregado  por pti1
    oDoc.WTextBox 670, 35, 90, 200, "", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
    oDoc.WTextBox 730, 35, 60, 180, "________________________________________", "F1", 8, hCenter
    oDoc.WTextBox 745, 90, 60, 80, "Firma", "F1", 10, hCenter
    
    sParrafoSeis = "¿Autorizas a Caja Maynas para el tratamiento de sus datos personales?"
    
    oDoc.WTextBox 670, 280, 60, 250, sParrafoSeis, "F1", 11, hLeft 'O  agregado  por pti1
   
   
    oDoc.WTextBox 712, 300, 15, 20, "SI", "F1", 8, hCenter
    oDoc.WTextBox 742, 300, 15, 20, "NO", "F1", 8, hCenter
    
    oDoc.WTextBox 690, 420, 70, 80, "", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
    oDoc.WTextBox 740, 280, 15, 20, "", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
    oDoc.WTextBox 710, 280, 15, 20, "", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
    If ultEstado = 1 Then
        oDoc.WTextBox 710, 280, 15, 20, "X", "F1", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, , 3
        oDoc.WTextBox 710, 280, 15, 20, "", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
    Else
        oDoc.WTextBox 740, 280, 15, 20, "", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
        oDoc.WTextBox 740, 280, 15, 20, "X", "F1", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, , 3
    End If
    

            
    oDoc.PDFClose
    oDoc.Show
    '</body>
End Sub



Private Function BuscaNombre(ByVal psNombre As String, ByVal nTipoBusqueda As TiposBusquedaNombre) As String 'add pti1 ers070-2018 19/12/2018
Dim sCadTmp As String
Dim PosIni As Integer
Dim PosFin As Integer
Dim PosIni2 As Integer
    sCadTmp = ""
    Select Case nTipoBusqueda
        Case 1 'Busqueda de Apellido Paterno
            If Mid(psNombre, 1, 1) <> "/" And Mid(psNombre, 1, 1) <> "\" And Mid(psNombre, 1, 1) <> "," Then
                PosIni = 1
                PosFin = InStr(1, psNombre, "/")
                If PosFin = 0 Then
                    PosFin = InStr(1, psNombre, "\")
                    If PosFin = 0 Then
                        PosFin = InStr(1, psNombre, ",")
                        If PosFin = 0 Then
                            PosFin = Len(psNombre)
                        End If
                    End If
                End If
                sCadTmp = Mid(psNombre, PosIni, PosFin - PosIni)
            Else
                sCadTmp = ""
            End If
        Case 2 'Apellido materno
           PosIni = InStr(1, psNombre, "/")
           If PosIni <> 0 Then
                PosIni = PosIni + 1
                PosFin = InStr(1, psNombre, "\")
                If PosFin = 0 Then
                    PosFin = InStr(1, psNombre, ",")
                    If PosFin = 0 Then
                        PosFin = Len(psNombre)
                    End If
                End If
                sCadTmp = Mid(psNombre, PosIni, PosFin - PosIni)
            Else
                sCadTmp = ""
            End If
        Case 3 'Apellido de casada
           PosIni = InStr(1, psNombre, "\")
           If PosIni <> 0 Then
                PosIni2 = InStr(1, psNombre, "VDA")
                If PosIni2 <> 0 Then
                    PosIni = PosIni2 + 3
                    PosFin = InStr(1, psNombre, ",")
                    If PosFin = 0 Then
                        PosFin = Len(psNombre)
                    End If
                Else
                    PosIni = PosIni + 1
                    PosFin = InStr(1, psNombre, ",")
                    If PosFin = 0 Then
                        PosFin = Len(psNombre)
                    End If
                End If
                sCadTmp = Trim(Mid(psNombre, PosIni, PosFin - PosIni))
            Else
                sCadTmp = ""
            End If
        Case 4 'Nombres
            PosIni = InStr(1, psNombre, ",")
            If PosIni <> 0 Then
                PosIni = PosIni + 1
                PosFin = Len(psNombre)
                sCadTmp = Mid(psNombre, PosIni, (PosFin + 1) - PosIni)
            Else
                sCadTmp = ""
            End If
            
    End Select
    BuscaNombre = sCadTmp
End Function
