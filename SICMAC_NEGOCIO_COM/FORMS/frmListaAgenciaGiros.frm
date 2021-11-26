VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmListaAgenciaGiros 
   Caption         =   "Ver..."
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3630
   Icon            =   "frmListaAgenciaGiros.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   3630
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   5106
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Agencias"
      TabPicture(0)   =   "frmListaAgenciaGiros.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lstAgencia"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.ListBox lstAgencia 
         Height          =   2400
         Left            =   75
         TabIndex        =   1
         Top             =   360
         Width           =   3490
      End
   End
End
Attribute VB_Name = "frmListaAgenciaGiros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre : frmListaAgenciaGiros
'** Descripción : Formulario creado con la finalida de listar las agencias seleccionadas para un determinado tarifario de giros
'** Creación : RECO, 20140410 - ERS008-2014
'**********************************************************************************************
Option Explicit

Public Sub Inicio(ByVal psGiroTarCod As String)
    Dim oServ As New COMNCaptaServicios.NCOMCaptaServicios
    Dim lrDatos As ADODB.Recordset
    Dim nIndex As Integer
    
    Set oServ = New COMNCaptaServicios.NCOMCaptaServicios
    Set lrDatos = New ADODB.Recordset
    
    Set lrDatos = oServ.ListaAgenciaTarGiros(psGiroTarCod)
    
    If Not (lrDatos.EOF And lrDatos.BOF) Then
        For nIndex = 0 To lrDatos.RecordCount - 1
            Me.lstAgencia.AddItem lrDatos!cAgeDescripcion
            lrDatos.MoveNext
        Next
    End If
    Me.Show 1
End Sub
