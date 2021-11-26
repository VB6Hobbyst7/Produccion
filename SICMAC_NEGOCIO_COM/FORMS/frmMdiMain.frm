VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDISicmact 
   BackColor       =   &H8000000C&
   ClientHeight    =   5820
   ClientLeft      =   1200
   ClientTop       =   2250
   ClientWidth     =   9480
   Icon            =   "frmMdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Tiempo 
      Interval        =   60000
      Left            =   0
      Top             =   0
   End
   Begin MSComctlLib.StatusBar SBBarra 
      Align           =   2  'Align Bottom
      Height          =   225
      Left            =   0
      TabIndex        =   0
      Top             =   5595
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   397
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   15875
            MinWidth        =   15875
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6174
            MinWidth        =   6174
         EndProperty
      EndProperty
   End
   Begin VB.Menu M0100000000 
      Caption         =   "&Archivo"
      Index           =   0
      Begin VB.Menu M0101000000 
         Caption         =   "Configurar &Impresora"
         Index           =   0
      End
      Begin VB.Menu M0101000000 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu M0101000000 
         Caption         =   "&Salir"
         Index           =   2
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu M0200000000 
      Caption         =   "&Captaciones"
      Index           =   0
      Begin VB.Menu M0201000000 
         Caption         =   "&Parámetros"
         Index           =   0
         Begin VB.Menu M0201010000 
            Caption         =   "&Mantenimiento"
            Index           =   0
         End
         Begin VB.Menu M0201010000 
            Caption         =   "&Consulta"
            Index           =   1
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "&Tasas de Interés"
         Index           =   1
         Begin VB.Menu M0201020000 
            Caption         =   "&Ahorros"
            Index           =   0
            Begin VB.Menu M0201020100 
               Caption         =   "&Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M0201020100 
               Caption         =   "&Consulta"
               Index           =   1
            End
         End
         Begin VB.Menu M0201020000 
            Caption         =   "&Plazo Fijo"
            Index           =   1
            Begin VB.Menu M0201020200 
               Caption         =   "&Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M0201020200 
               Caption         =   "&Consulta"
               Index           =   1
            End
         End
         Begin VB.Menu M0201020000 
            Caption         =   "&CTS"
            Index           =   2
            Begin VB.Menu M0201020300 
               Caption         =   "&Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M0201020300 
               Caption         =   "&Consulta"
               Index           =   1
            End
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "&Mantenimiento"
         Index           =   2
         Begin VB.Menu M0201030000 
            Caption         =   "&Ahorros"
            Index           =   0
         End
         Begin VB.Menu M0201030000 
            Caption         =   "&Plazo Fijo"
            Index           =   1
         End
         Begin VB.Menu M0201030000 
            Caption         =   "&CTS"
            Index           =   2
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "&Bloqueo/Desbloqueo"
         Index           =   3
         Begin VB.Menu M0201040000 
            Caption         =   "&Ahorros"
            Index           =   0
         End
         Begin VB.Menu M0201040000 
            Caption         =   "&Plazo Fijo"
            Index           =   1
         End
         Begin VB.Menu M0201040000 
            Caption         =   "&CTS"
            Index           =   2
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "&Simulación Plazo Fijo"
         Index           =   4
      End
      Begin VB.Menu M0201000000 
         Caption         =   "&Tarjeta Magnética"
         Index           =   5
         Begin VB.Menu M0201050000 
            Caption         =   "&Registro"
            Index           =   0
         End
         Begin VB.Menu M0201050000 
            Caption         =   "Re&lación"
            Index           =   1
         End
         Begin VB.Menu M0201050000 
            Caption         =   "&Bloqueo/Desbloqueo"
            Index           =   2
         End
         Begin VB.Menu M0201050000 
            Caption         =   "&Cambio de Clave"
            Index           =   3
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "Bene&ficiarios"
         Index           =   6
         Begin VB.Menu M0201060000 
            Caption         =   "&Mantenimiento"
            Index           =   0
         End
         Begin VB.Menu M0201060000 
            Caption         =   "&Consulta "
            Index           =   1
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "&Orden Pago"
         Index           =   7
         Begin VB.Menu M0201070000 
            Caption         =   "&Generación y Emisión"
            Index           =   0
         End
         Begin VB.Menu M0201070000 
            Caption         =   "&Certificación"
            Index           =   1
         End
         Begin VB.Menu M0201070000 
            Caption         =   "&Anulación"
            Index           =   2
         End
         Begin VB.Menu M0201070000 
            Caption         =   "Con&sulta"
            Index           =   3
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "Ser&vicios"
         Index           =   8
         Begin VB.Menu M0201080000 
            Caption         =   "&Sedalib"
            Index           =   0
            Begin VB.Menu M0201080100 
               Caption         =   "&Parámetros"
               Index           =   0
            End
            Begin VB.Menu M0201080100 
               Caption         =   "&Reportes"
               Index           =   1
            End
            Begin VB.Menu M0201080100 
               Caption         =   "&Generación DBF"
               Index           =   2
            End
         End
         Begin VB.Menu M0201080000 
            Caption         =   "&Hidrandina"
            Index           =   1
            Begin VB.Menu M0201080200 
               Caption         =   "&Parámetros"
               Index           =   0
            End
            Begin VB.Menu M0201080200 
               Caption         =   "&Reportes"
               Index           =   1
            End
         End
         Begin VB.Menu M0201080000 
            Caption         =   "&Universidad Nacional de Trujillo"
            Index           =   2
            Begin VB.Menu M0201080300 
               Caption         =   "&Parámetros"
               Index           =   0
            End
            Begin VB.Menu M0201080300 
               Caption         =   "&Reportes"
               Index           =   1
            End
            Begin VB.Menu M0201080300 
               Caption         =   "&Generación DBF"
               Index           =   2
            End
         End
         Begin VB.Menu M0201080000 
            Caption         =   "&Instituciones Convenio"
            Index           =   3
            Begin VB.Menu M0201080400 
               Caption         =   "&Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M0201080400 
               Caption         =   "&Cuentas"
               Index           =   1
            End
            Begin VB.Menu M0201080400 
               Caption         =   "&Plan de Pagos"
               Index           =   2
            End
         End
      End
   End
   Begin VB.Menu M0300000000 
      Caption         =   "Colocacio&nes"
      Index           =   0
      Begin VB.Menu M0301000000 
         Caption         =   "Cartas &Fianza"
         Index           =   0
         Begin VB.Menu M0301010000 
            Caption         =   "&Solicitud Carta Fianza"
            Index           =   0
         End
         Begin VB.Menu M0301010000 
            Caption         =   "Gra&var Garantia"
            Index           =   1
         End
         Begin VB.Menu M0301010000 
            Caption         =   "Sugerencia de &Analista"
            Index           =   2
         End
         Begin VB.Menu M0301010000 
            Caption         =   "&Resolver Carta Fianza"
            Index           =   3
            Begin VB.Menu M0301010100 
               Caption         =   "&Aprobar Carta Fianza"
               Index           =   0
            End
            Begin VB.Menu M0301010100 
               Caption         =   "&Rechazar Carta Fianza"
               Index           =   1
            End
         End
         Begin VB.Menu M0301010000 
            Caption         =   "&Emitir Carta Fianza"
            Index           =   4
         End
         Begin VB.Menu M0301010000 
            Caption         =   "&Honrar Carta Fianza"
            Index           =   5
         End
         Begin VB.Menu M0301010000 
            Caption         =   "Retirar Carta Fianza Aprobada"
            Index           =   6
         End
         Begin VB.Menu M0301010000 
            Caption         =   "Niveles de Aprobacion"
            Index           =   7
         End
         Begin VB.Menu M0301010000 
            Caption         =   "Reportes"
            Index           =   8
         End
      End
      Begin VB.Menu M0301000000 
         Caption         =   "&Creditos"
         Index           =   1
         Begin VB.Menu M0301020000 
            Caption         =   "&Definiciones"
            Index           =   0
            Begin VB.Menu M0301020100 
               Caption         =   "&Parametros de Control"
               Index           =   0
               Begin VB.Menu M0301020101 
                  Caption         =   "&Mantenimiento"
                  Index           =   0
               End
               Begin VB.Menu M0301020101 
                  Caption         =   "&Consulta"
                  Index           =   1
               End
            End
            Begin VB.Menu M0301020100 
               Caption         =   "&Lineas de Credito"
               Index           =   1
               Begin VB.Menu M0301020102 
                  Caption         =   "&Registro"
                  Index           =   0
               End
               Begin VB.Menu M0301020102 
                  Caption         =   "&Mantenimiento"
                  Index           =   1
               End
               Begin VB.Menu M0301020102 
                  Caption         =   "&Consulta"
                  Index           =   2
               End
            End
            Begin VB.Menu M0301020100 
               Caption         =   "&Niveles de Aprobacion"
               Index           =   2
               Begin VB.Menu M0301020103 
                  Caption         =   "&Mantenimiento"
                  Index           =   0
               End
               Begin VB.Menu M0301020103 
                  Caption         =   "&Consulta"
                  Index           =   1
               End
            End
            Begin VB.Menu M0301020100 
               Caption         =   "&Gastos"
               Index           =   3
               Begin VB.Menu M0301020104 
                  Caption         =   "&Mantenimiento"
                  Index           =   0
               End
               Begin VB.Menu M0301020104 
                  Caption         =   "&Consulta"
                  Index           =   1
               End
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "&Solicitud Credito"
            Index           =   1
            Begin VB.Menu M0301020200 
               Caption         =   "&Registro"
               Index           =   0
            End
            Begin VB.Menu M0301020200 
               Caption         =   "&Consulta"
               Index           =   1
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Relaciones de Credito"
            Index           =   2
            Begin VB.Menu M0301020300 
               Caption         =   "&Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M0301020300 
               Caption         =   "Reasignacion de &Cartera en Lote"
               Index           =   1
            End
            Begin VB.Menu M0301020300 
               Caption         =   "Con&sulta"
               Index           =   2
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "&Garantias"
            Index           =   3
            Begin VB.Menu M0301020400 
               Caption         =   "&Registro"
               Index           =   0
            End
            Begin VB.Menu M0301020400 
               Caption         =   "&Mantenimiento"
               Index           =   1
            End
            Begin VB.Menu M0301020400 
               Caption         =   "&Consulta"
               Index           =   2
            End
            Begin VB.Menu M0301020400 
               Caption         =   "Gravamen"
               Index           =   3
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "&Sugerencia"
            Index           =   4
            Begin VB.Menu M0301020500 
               Caption         =   "&Registro"
               Index           =   0
            End
            Begin VB.Menu M0301020500 
               Caption         =   "&Registro"
               Index           =   1
            End
            Begin VB.Menu M0301020500 
               Caption         =   "&Consulta"
               Index           =   2
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "&Resolver Creditos"
            Index           =   5
            Begin VB.Menu M0301020600 
               Caption         =   "&Aprobacion"
               Index           =   0
            End
            Begin VB.Menu M0301020600 
               Caption         =   "&Rechazo"
               Index           =   1
            End
            Begin VB.Menu M0301020600 
               Caption         =   "A&nulacion Creditos Aprobados"
               Index           =   2
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "&Reprogramacion Credito"
            Index           =   6
            Begin VB.Menu M0301020700 
               Caption         =   "Repr&ogramacion"
               Index           =   0
            End
            Begin VB.Menu M0301020700 
               Caption         =   "Reprogramacion en &Lote"
               Index           =   1
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "&Refinanciacion"
            Index           =   7
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Actualizacion de &Metodos de Liquidacion"
            Index           =   8
         End
         Begin VB.Menu M0301020000 
            Caption         =   "&Perdonar Mora"
            Index           =   9
         End
         Begin VB.Menu M0301020000 
            Caption         =   "&Gastos"
            Index           =   10
            Begin VB.Menu M0301020800 
               Caption         =   "&Administracion de Gastos en &Lote"
               Index           =   0
            End
            Begin VB.Menu M0301020800 
               Caption         =   "Mantenimiento de Penalidad de Cancelacion"
               Index           =   1
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Reasignar &Institucion"
            Index           =   11
         End
         Begin VB.Menu M0301020000 
            Caption         =   "&Transferencia a Recuperaciones"
            Index           =   12
         End
         Begin VB.Menu M0301020000 
            Caption         =   "&Analista"
            Index           =   13
            Begin VB.Menu M0301020900 
               Caption         =   "&Nota de Analista"
               Index           =   0
            End
            Begin VB.Menu M0301020900 
               Caption         =   "&Metas de Analista"
               Index           =   1
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Si&mulaciones"
            Index           =   14
            Begin VB.Menu M0301021000 
               Caption         =   "Calendario de &Pagos"
               Index           =   0
            End
            Begin VB.Menu M0301021000 
               Caption         =   "Calendario de &Desembolsos Parciales"
               Index           =   1
            End
            Begin VB.Menu M0301021000 
               Caption         =   "Calendario de Cuota &Libre"
               Index           =   2
            End
            Begin VB.Menu M0301021000 
               Caption         =   "Simulador de Pagos"
               Index           =   3
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Cons&ultas"
            Index           =   15
            Begin VB.Menu M0301021100 
               Caption         =   "&Historial de Credito"
               Index           =   0
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "&Reportes"
            Index           =   16
            Begin VB.Menu M0301021200 
               Caption         =   "&Duplicados"
               Index           =   0
            End
            Begin VB.Menu M0301021200 
               Caption         =   "Listados de Creditos"
               Index           =   1
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Registro de Dacion de Pago"
            Index           =   17
         End
      End
      Begin VB.Menu M0301000000 
         Caption         =   "&Pignoraticio"
         Index           =   2
         Begin VB.Menu M0301030000 
            Caption         =   "&Contratro"
            Index           =   0
            Begin VB.Menu M0301030100 
               Caption         =   "&Registro"
               Index           =   0
            End
            Begin VB.Menu M0301030100 
               Caption         =   "Mantenimiento Descripcion"
               Index           =   1
            End
            Begin VB.Menu M0301030100 
               Caption         =   "Anulación"
               Index           =   2
            End
            Begin VB.Menu M0301030100 
               Caption         =   "&Bloqueo"
               Index           =   3
            End
         End
         Begin VB.Menu M0301030000 
            Caption         =   "Rescate Joyas"
            Index           =   1
         End
         Begin VB.Menu M0301030000 
            Caption         =   "Remate"
            Index           =   2
            Begin VB.Menu M0301030200 
               Caption         =   "Preparacion Remate"
               Index           =   0
            End
            Begin VB.Menu M0301030200 
               Caption         =   "Remate"
               Index           =   1
            End
         End
         Begin VB.Menu M0301030000 
            Caption         =   "Adjudicacion"
            Index           =   3
         End
         Begin VB.Menu M0301030000 
            Caption         =   "Subasta"
            Index           =   4
            Begin VB.Menu M0301030300 
               Caption         =   "Preparacion Subasta"
               Index           =   0
            End
            Begin VB.Menu M0301030300 
               Caption         =   "Subasta"
               Index           =   1
            End
         End
         Begin VB.Menu M0301030000 
            Caption         =   "&Reportes"
            Index           =   5
            Begin VB.Menu M0301030400 
               Caption         =   "Movimientos Diario"
               Index           =   0
            End
            Begin VB.Menu M0301030400 
               Caption         =   "Listados Generales"
               Index           =   1
            End
            Begin VB.Menu M0301030400 
               Caption         =   "Estadisticas"
               Index           =   2
            End
            Begin VB.Menu M0301030400 
               Caption         =   "Contrato de Persona"
               Index           =   3
            End
         End
      End
      Begin VB.Menu M0301000000 
         Caption         =   "&Recuperaciones"
         Index           =   3
         Begin VB.Menu M0301040000 
            Caption         =   "&Comisiones Abogados"
            Index           =   0
         End
         Begin VB.Menu M0301040000 
            Caption         =   "&Registro de Otros Creditos "
            Index           =   1
         End
         Begin VB.Menu M0301040000 
            Caption         =   "Proceso &Judicial"
            Index           =   2
            Begin VB.Menu M0301040100 
               Caption         =   "&Expedientes Judiciales"
               Index           =   0
            End
            Begin VB.Menu M0301040100 
               Caption         =   "Actuaciones &Procesales"
               Index           =   1
            End
         End
         Begin VB.Menu M0301040000 
            Caption         =   "&Gastos de Recuperacines"
            Index           =   3
         End
         Begin VB.Menu M0301040000 
            Caption         =   "Metodo de &Liquidacion"
            Index           =   4
         End
         Begin VB.Menu M0301040000 
            Caption         =   "Cas&tigar Credito"
            Index           =   5
         End
         Begin VB.Menu M0301040000 
            Caption         =   "Pago de Creditos"
            Index           =   6
         End
         Begin VB.Menu M0301040000 
            Caption         =   "Extorno de Pagos"
            Index           =   7
         End
         Begin VB.Menu M0301040000 
            Caption         =   "Consulta Para&metros"
            Index           =   8
         End
         Begin VB.Menu M0301040000 
            Caption         =   "Mantenimiento Parametros"
            Index           =   9
         End
      End
   End
   Begin VB.Menu M0400000000 
      Caption         =   "&Operaciones"
      Index           =   0
      Begin VB.Menu M0401000000 
         Caption         =   "Tipo de Ca&mbio"
         Index           =   0
      End
      Begin VB.Menu M0401000000 
         Caption         =   "&Habilitación/Devolución"
         Index           =   1
         Begin VB.Menu M0401010000 
            Caption         =   "&Cajero"
            Index           =   0
         End
         Begin VB.Menu M0401010000 
            Caption         =   "&Administracion"
            Index           =   1
         End
      End
      Begin VB.Menu M0401000000 
         Caption         =   "&Billetaje"
         Index           =   2
      End
      Begin VB.Menu M0401000000 
         Caption         =   "&Operaciones"
         Index           =   3
         Shortcut        =   {F2}
      End
      Begin VB.Menu M0401000000 
         Caption         =   "Operaciones CMACs &Recepción"
         Index           =   4
         Shortcut        =   {F3}
      End
      Begin VB.Menu M0401000000 
         Caption         =   "Operaciones CMACs &Llamada"
         Index           =   5
         Shortcut        =   {F4}
      End
      Begin VB.Menu M0401000000 
         Caption         =   "&Estadisticas de Compra Venta de ME"
         Index           =   6
      End
      Begin VB.Menu M0401000000 
         Caption         =   "Tarjeta &Magnetica"
         Index           =   7
         Begin VB.Menu M0401020000 
            Caption         =   "&Registro de Tarjeta"
            Index           =   0
         End
         Begin VB.Menu M0401020000 
            Caption         =   "Mantenimiento de Tarjetas Magneticas"
            Index           =   1
         End
      End
      Begin VB.Menu M0401000000 
         Caption         =   "&Cierres"
         Index           =   8
         Begin VB.Menu M0401030000 
            Caption         =   "Cierre de &Día"
            Index           =   0
         End
         Begin VB.Menu M0401030000 
            Caption         =   "Cierre de &Mes"
            Index           =   1
         End
         Begin VB.Menu M0401030000 
            Caption         =   "Cierre Diario Batch"
            Index           =   2
         End
         Begin VB.Menu M0401030000 
            Caption         =   "Cierre Mes Batch"
            Index           =   3
         End
      End
      Begin VB.Menu M0401000000 
         Caption         =   "E&xtornos "
         Index           =   9
         Begin VB.Menu M0401040000 
            Caption         =   "Extornos &Ahorros"
            Index           =   0
         End
         Begin VB.Menu M0401040000 
            Caption         =   "Extornos &Créditos"
            Index           =   1
            Begin VB.Menu M0401040100 
               Caption         =   "Extornos Desembolsos"
               Index           =   0
            End
            Begin VB.Menu M0401040100 
               Caption         =   "Extornos Pagos"
               Index           =   1
            End
         End
         Begin VB.Menu M0401040000 
            Caption         =   "Extornos &Pignoraticio"
            Index           =   2
         End
         Begin VB.Menu M0401040000 
            Caption         =   "Extornos &Judicial"
            Index           =   3
         End
      End
      Begin VB.Menu M0401000000 
         Caption         =   "Asiento Contable"
         Index           =   10
         Begin VB.Menu M0401050000 
            Caption         =   "Asiento Contable &Dia"
            Index           =   0
         End
         Begin VB.Menu M0401050000 
            Caption         =   "Asiento Contable &Anterior"
            Index           =   1
         End
         Begin VB.Menu M0401050000 
            Caption         =   "Mantenimiento &Asiento Contable"
            Index           =   2
         End
      End
      Begin VB.Menu M0401000000 
         Caption         =   "&Reportes"
         Index           =   11
         Begin VB.Menu M0401060000 
            Caption         =   "&Resumen Ingresos y Egresos"
            Index           =   0
         End
         Begin VB.Menu M0401060000 
            Caption         =   "Resumen de Ingresos y Egresos Consolidado"
            Index           =   1
         End
         Begin VB.Menu M0401060000 
            Caption         =   "Detalle de &Operaciones"
            Index           =   2
         End
         Begin VB.Menu M0401060000 
            Caption         =   "Detalle de &Habilitación/Devolución Cajero"
            Index           =   3
         End
         Begin VB.Menu M0401060000 
            Caption         =   "Protocolo por &Usuario"
            Index           =   4
         End
         Begin VB.Menu M0401060000 
            Caption         =   "&Total Operaciones de Usuario"
            Index           =   5
         End
         Begin VB.Menu M0401060000 
            Caption         =   "Reportes de Cajamatic"
            Index           =   6
         End
      End
   End
   Begin VB.Menu M0500000000 
      Caption         =   "&Servicios"
      Index           =   0
      Begin VB.Menu M0501000000 
         Caption         =   "U.N.T"
         Index           =   0
         Begin VB.Menu M0501010000 
            Caption         =   "&Registrar Alumnos Nuevos UNT"
            Index           =   0
         End
         Begin VB.Menu M0501010000 
            Caption         =   "Generación Archivo &UNT"
            Index           =   1
         End
         Begin VB.Menu M0501010000 
            Caption         =   "&Actualiza Orden PAgo UNT"
            Index           =   2
         End
      End
      Begin VB.Menu M0501000000 
         Caption         =   "Telefonica"
         Index           =   1
         Begin VB.Menu M0501020000 
            Caption         =   "&Generación Archivo Telefonica"
            Index           =   0
         End
      End
      Begin VB.Menu M0501000000 
         Caption         =   "Hidrandina/Sedalib"
         Index           =   2
         Begin VB.Menu M0501030000 
            Caption         =   "Consolidacion de Ser&vicios"
            Index           =   0
         End
      End
   End
   Begin VB.Menu M0600000000 
      Caption         =   "&Sistema"
      Index           =   0
      Begin VB.Menu M0601000000 
         Caption         =   "&Parametros del Sistema"
         Index           =   0
      End
      Begin VB.Menu M0601000000 
         Caption         =   "&Variables del Sistema"
         Index           =   1
      End
      Begin VB.Menu M0601000000 
         Caption         =   "Mantenimiento &Permisos"
         Index           =   4
      End
      Begin VB.Menu M0601000000 
         Caption         =   "Mantenimiento &Zonas"
         Index           =   6
      End
      Begin VB.Menu M0601000000 
         Caption         =   "Mantenimiento Agencias"
         Index           =   7
      End
      Begin VB.Menu M0601000000 
         Caption         =   "Copia de &Seguridad"
         Index           =   9
      End
   End
   Begin VB.Menu M0700000000 
      Caption         =   "&Personas"
      Index           =   0
      Begin VB.Menu M0701000000 
         Caption         =   "&Personas"
         Index           =   0
         Begin VB.Menu M0701010000 
            Caption         =   "&Registro"
            Index           =   0
         End
         Begin VB.Menu M0701010000 
            Caption         =   "&Manteniemiento"
            Index           =   1
         End
         Begin VB.Menu M0701010000 
            Caption         =   "&Consulta"
            Index           =   2
         End
         Begin VB.Menu M0701010000 
            Caption         =   "&Exoneradas del Lavado de Dinero"
            Index           =   3
         End
      End
      Begin VB.Menu M0701000000 
         Caption         =   "&Instituciones Financieras"
         Index           =   1
         Begin VB.Menu M0701020000 
            Caption         =   "Mantenimiento"
            Index           =   0
         End
         Begin VB.Menu M0701020000 
            Caption         =   "Consulta"
            Index           =   1
         End
      End
      Begin VB.Menu M0701000000 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu M0701000000 
         Caption         =   "&Posicion de Cliente"
         Index           =   3
      End
      Begin VB.Menu M0701000000 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu M0701000000 
         Caption         =   "Central de Riesgos"
         Index           =   5
         Begin VB.Menu M0701030000 
            Caption         =   "&SuperIntendencia de Banca y Seguros"
            Index           =   0
            Shortcut        =   ^{F1}
         End
         Begin VB.Menu M0701030000 
            Caption         =   "&Infocorp"
            Index           =   1
            Shortcut        =   ^{F2}
         End
         Begin VB.Menu M0701030000 
            Caption         =   "&Camara de Comercio"
            Index           =   2
            Shortcut        =   ^{F3}
         End
         Begin VB.Menu M0701030000 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu M0701030000 
            Caption         =   "Morosos"
            Index           =   4
            Shortcut        =   ^{F4}
         End
         Begin VB.Menu M0701030000 
            Caption         =   "Direll"
            Index           =   5
            Shortcut        =   ^{F5}
         End
      End
   End
   Begin VB.Menu M0800000000 
      Caption         =   "Herra&mientas"
      Index           =   0
      Begin VB.Menu M0801000000 
         Caption         =   "Editor de Textos"
         Index           =   0
      End
      Begin VB.Menu M0801000000 
         Caption         =   "Spooler de Impresión"
         Index           =   1
      End
      Begin VB.Menu M0801000000 
         Caption         =   "Configuración de Periféricos"
         Index           =   2
      End
      Begin VB.Menu M0801000000 
         Caption         =   "Cargar Logo &Penware"
         Index           =   3
      End
      Begin VB.Menu M0801000000 
         Caption         =   "Explorador de Archivos"
         Index           =   4
      End
   End
   Begin VB.Menu M0900000000 
      Caption         =   "A&yuda"
      Index           =   0
      Begin VB.Menu M0901000000 
         Caption         =   "&Contenido"
         Index           =   0
         Shortcut        =   {F1}
      End
      Begin VB.Menu M0901000000 
         Caption         =   "&Indice"
         Index           =   1
      End
      Begin VB.Menu M0901000000 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu M0901000000 
         Caption         =   "&Acerca del Sistema..."
         Index           =   3
      End
   End
End
Attribute VB_Name = "MDISicmact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Api para Ejecutar Internet Explorer
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
"ShellExecuteA" (ByVal hWnd As Long, ByVal lpoperation As String, _
ByVal lpfile As String, ByVal lpParameters As String, _
ByVal lpDirectory As String, ByVal nShowcmd As Long) As Long

Private Sub M0101000000_Click(Index As Integer)
    Select Case Index
        Case 0
            frmImpresora.Show 1
        Case 2
            If MsgBox("Esta Seguro que Desea Salir del Sistema ?", vbInformation + vbYesNo, "Aviso") = vbYes Then
                End
            End If
    End Select
End Sub

Private Sub M0201000000_Click(Index As Integer)
    Select Case Index
        Case 4
            frmCapSimulacionPF.Show 1
    End Select
End Sub

Private Sub M0201010000_Click(Index As Integer)
    Select Case Index
        Case 0
             frmCapParametros.Inicia False
        Case 1
            frmCapParametros.Inicia True
    End Select
End Sub

Private Sub M0201020100_Click(Index As Integer)
    Select Case Index
        Case 0 'mantenimiento
            frmCapTasaInt.Inicia gCapAhorros, False
        Case 1 'Consulta
            frmCapTasaInt.Inicia gCapAhorros, True
    End Select
End Sub

Private Sub M0201020200_Click(Index As Integer)
    Select Case Index
        Case 0 'Mantenimiento
            frmCapTasaInt.Inicia gCapPlazoFijo, False
        Case 1 'Consulta
            frmCapTasaInt.Inicia gCapPlazoFijo, True
    End Select
End Sub

Private Sub M0201020300_Click(Index As Integer)
    Select Case Index
        Case 0
            frmCapTasaInt.Inicia gCapCTS, True
        Case 1
            frmCapTasaInt.Inicia gCapCTS, False
    End Select
End Sub

Private Sub M0201030000_Click(Index As Integer)
    'Mantenimiento
    Select Case Index
        Case 0 'Ahorro
            frmCapMantenimiento.Inicia gCapAhorros
        Case 1 'Plazo fijo
            frmCapMantenimiento.Inicia gCapPlazoFijo
        Case 2 'Cts
            frmCapMantenimiento.Inicia gCapCTS
    End Select
End Sub

Private Sub M0201040000_Click(Index As Integer)
    'Bloqueos / Desbloqueos
    Select Case Index
        Case 0 'Ahorros
            frmCapBloqueoDesbloqueo.Inicia gCapAhorros
        Case 1 'Plazo Fijo
            frmCapBloqueoDesbloqueo.Inicia gCapPlazoFijo
        Case 2 'Cts
            frmCapBloqueoDesbloqueo.Inicia gCapCTS
    End Select
End Sub

Private Sub M0201050000_Click(Index As Integer)
    Select Case Index
        Case 0 'Registro
            frmCapTarjetaRegistro.Inicia
        Case 1 'Relacion
            frmCapTarjetaRelacion.Inicia False
        Case 2 'Bloqueo
            frmCapTarjetaBlqDesBlq.Show 1
        Case 3 'Cambio de Clave
            frmCapTarjetaCambioClave.Show 1
    End Select
End Sub

Private Sub M0201060000_Click(Index As Integer)
    Select Case Index
        Case 0
            frmCapBeneficiario.Inicia False
        Case 1
            frmCapBeneficiario.Inicia True
    End Select
End Sub

Private Sub M0201070000_Click(Index As Integer)
    Select Case Index
        Case 0 'generacion
            frmCapOrdPagGenEmi.Show 1
        Case 1 'Certificacion
            frmCapOrdPagAnulCert.Inicia gAhoOPCertificacion
        Case 2 'Anulacion
            frmCapOrdPagAnulCert.Inicia gAhoOPAnulacion
        Case 3 'Consulta
            frmCapOrdPagConsulta.Show 1
    End Select
End Sub

Private Sub M0201080100_Click(Index As Integer)
    Select Case Index
        Case 0 'Parametros
            frmCapServParametros.Inicia gCapServSedalib
        Case 1 'reporte
            frmCapServGenReporte.Inicia gCapServSedalib
        Case 2 'Generacion de DBF
            frmCapServGeneraDBF.Show 1
    End Select
End Sub

Private Sub M0201080200_Click(Index As Integer)
    Select Case Index
        Case 0 'parametros
            frmCapServParametros.Inicia gCapServHidrandina
        Case 1 'reportes
            frmCapServGenReporte.Inicia gCapServHidrandina
    End Select
End Sub

Private Sub M0201080300_Click(Index As Integer)

End Sub

Private Sub M0201080400_Click(Index As Integer)
    Select Case Index
        Case 0 'mantenimiento
            frmCapConvenioMant.Show 1
        Case 1 'Cuentas
            frmCapServConvCuentas.Inicia
        Case 2 'Plan Pagos
            frmCapServConvPlanPag.Inicia
    End Select
End Sub

Private Sub M0301020000_Click(Index As Integer)
    Select Case Index
    Case 7 'Refinanciacion de Credito
        Call frmCredSolicitud.RefinanciaCredito(Registrar)
    Case 8 ' 'Actualizacion con Metodos de Liquidacion
        frmCredMntMetLiquid.Show 1
    Case 9 'Perdonar Mora
        frmCredPerdonarMora.Show 1
    Case 11 'Reasignar Institucion
        frmCredReasigInst.Show 1
    Case 12 'Transferencia a Recuperaciones
        frmCredTransARecup.Show 1
    Case 17 'Registro de Dacion de Pago
        frmCredRegisDacion.Show 1
    End Select
End Sub

Private Sub M0301020101_Click(Index As Integer)
    Select Case Index
        Case 0 'Mantenimeinto de Parametros
            frmCredMantParametros.InicioActualizar
        Case 1 'Consulta de Parametros
            frmCredMantParametros.InicioCosultar
    End Select
End Sub

Private Sub M0301020102_Click(Index As Integer)
    Select Case Index
        Case 0 'Registro de Lineas de Credito
            frmCredLineaCredito.Registrar
        Case 1 'Mantenimiento de Lineas de Credito
            frmCredLineaCredito.Actualizar
        Case 2 ' Consulta de lineas de Credito
            frmCredLineaCredito.Consultar
    End Select
End Sub

Private Sub M0301020103_Click(Index As Integer)
    Select Case Index
        Case 0 'Mantenimeinto de Niveles de Aprobacion
            frmCredNivAprCred.Inicio MuestraNivelesActualizar
        Case 1 'Consulta de Niveles de Aprobacion
            frmCredNivAprCred.Inicio MuestraNivelesConsulta
    End Select
End Sub

Private Sub M0301020104_Click(Index As Integer)
    Select Case Index
        Case 0 'Mantenimeinto de Gastos
            frmCredMntGastos.Inicio InicioGastosActualizar
        Case 1
            frmCredMntGastos.Inicio InicioGastosConsultar
    End Select
End Sub

Private Sub M0301020200_Click(Index As Integer)
    Select Case Index
        Case 0 'Registro de Solicitud
            frmCredSolicitud.Inicio Registrar
        Case 1 'Consulta de Solicitud
            frmCredSolicitud.Inicio Consulta
    End Select
End Sub

Private Sub M0301020300_Click(Index As Integer)
    Dim oCredRel As New UCredRelacion

    Select Case Index
        Case 0 'Mantenimiento de Relaciones de Credito
            frmCredRelaCta.Inicio oCredRel, InicioMantenimiento, InicioRegistroForm
            Set oCredRel = Nothing
        Case 1 'Reasignacion de Cartera en Lote
            frmCredReasigCartera.Show 1
        Case 2 'Consulta de Relaciones de Credito
            frmCredRelaCta.Inicio oCredRel, InicioMantenimiento, InicioConsultaForm
            Set oCredRel = Nothing
    End Select
End Sub

Private Sub M0301020400_Click(Index As Integer)
    Select Case Index
        Case 0 'Registro de Garantia
            frmPersGarantias.Inicio RegistroGarantia
        Case 1 'Mantenimiento de Garantia
            frmPersGarantias.Inicio MantenimientoGarantia
        Case 2 'Consulta de Garantia
            frmPersGarantias.Inicio ConsultaGarant
        Case 3 'Gravament
            frmCredGarantCred.Inicio PorMenu
    End Select
End Sub

Private Sub M0301020500_Click(Index As Integer)
    Select Case Index
        Case 0 'Registro de Sugerencia
            frmCredSugerencia.Show 1
        Case 1 'Mantenimiento de Sugerencia
            
        Case 2 'Consulta de Sugerencia
            
    End Select
End Sub

Private Sub M0301020600_Click(Index As Integer)
    Select Case Index
        Case 0 'Aprobacion de Credito
            frmCredAprobacion.Show 1
        Case 1 'Rechazo de Credito
            frmCredRechazo.Rechazar
        Case 2 'Anulacion de Credito
            frmCredRechazo.Retirar
    End Select
End Sub

Private Sub M0301020700_Click(Index As Integer)
    Select Case Index
        Case 0 'Reprogramacion de Credito
            frmCredReprogCred.Show 1
        Case 1 'Reprogramacion en Lote
            frmCredReprogLote.Show 1
    End Select
End Sub

Private Sub M0301020800_Click(Index As Integer)
    Select Case Index
        Case 0 'Administracion de Gastos en Lote
            frmCredAsigGastosLote.Show 1
        Case 1 ' mantenimiento de Penalidad
            frmCredExonerarPen.Show 1
    End Select
End Sub

Private Sub M0301020900_Click(Index As Integer)
    Select Case Index
        Case 0 'Nota del Analista
            frmCredAsigNota.Show 1
        Case 1 'Meta del Analista
            frmCredMetasAnalista.Show 1
    End Select
End Sub

Private Sub M0301021000_Click(Index As Integer)
    Dim MatCalend As Variant
    Dim Matriz(0) As String
    
    Select Case Index
        Case 0 'Calendario de Pagos
            frmCredCalendPagos.Simulacion DesembolsoTotal
        Case 1 'Desembolsos Parciales
            frmCredCalendPagos.Simulacion DesembolsoParcial
        Case 2 'Cuota Libre
            MatCalend = frmCredCalendCuotaLibre.CalendarioLibre(True, gdFecSis, Matriz, 0#, 0, 0#)
    End Select
End Sub

Private Sub M0301021100_Click(Index As Integer)
    If Index = 0 Then
        frmCredConsulta.Show 1
    End If
End Sub

Private Sub M0301021200_Click(Index As Integer)
    Select Case Index
        Case 0
            frmCredDupDoc.Show 1
        Case 1
            frmCredReportes.Show 1
    End Select
End Sub

Private Sub M0301030000_Click(Index As Integer)
    Select Case Index
        Case 2
            frmColPRescateJoyas.Show 1
        Case 4
            'frmColPAdjudicaLotes.Show 1
    End Select
End Sub

Private Sub M0301030100_Click(Index As Integer)
    Select Case Index
        Case 0 'Registro
            frmColPRegContrato.Show 1
        Case 1 'Mantenimiento
            frmColPMantPrestamoPig.Show 1
        Case 2 'Anulacion
            frmColPAnularPrestamoPig.Show 1
        Case 3 'Bloqueo
            frmColPBloqueo.Show 1
    End Select
End Sub

Private Sub M0301030200_Click(Index As Integer)
    Select Case Index
        Case 0 'Preparacion de Remate
            frmColPRematePrepara.Show 1
        Case 1 'Remate
            frmColPRemateProceso.Show 1
    End Select
End Sub

Private Sub M0301030300_Click(Index As Integer)
    Select Case Index
        Case 0
            frmColPSubastaPrepara.Show 1
        Case 1
            frmColPSubastaProceso.Show 1
    End Select
End Sub

Private Sub M0301030400_Click(Index As Integer)
    Select Case Index
        Case 0
            
        Case 1
            
        Case 2
            
        Case 3
            
    End Select
End Sub

Private Sub M0301040000_Click(Index As Integer)
    Select Case Index
        Case 0
            frmColRecComision.Show 1
        Case 1
            frmColRecIngresoOtrasEnt.Show 1
        Case 2
            
        Case 3
            frmColRecGastosRecuperaciones.Show 1
        Case 4
            frmColRecMetodoLiquid.Show 1
        Case 5
        
        Case 6
            frmColRecPagoCredRecup.Show 1
        Case 7
        Case 8
        Case 9
            frmColocEvalCalCli.Show 1
    End Select
End Sub

Private Sub M0301040100_Click(Index As Integer)
    Select Case Index
        Case 0
            frmColRecExped.Show 1
        Case 1
        
    End Select
End Sub

Private Sub M0401000000_Click(Index As Integer)
    Select Case Index
        Case 0
            frmMantTipoCambio.Show 1
        Case 1
            
        Case 2
            
        Case 3
            frmCajeroOperaciones.Show 1
        Case 5
            frmCajeroOpeCMAC.Inicia
        Case 6
            frmCajeroOpeCMAC.Inicia False
            
    End Select
End Sub

Private Sub M0401030000_Click(Index As Integer)
    Select Case Index
        Case 0
            Call frmCierreDiario.CierreDia
        Case 1
            Call frmCierreDiario.CierreMes
        Case 2
        
        Case 3
            
    End Select
            
End Sub

Private Sub M0401040000_Click(Index As Integer)
    Select Case Index
        Case 0 'Captaciones
            frmCapExtornos.Show 1
        
            
        Case 1 '    Extorno Credito
        
        Case 2 '    Pignoraticio
        
        Case 3 '    recuperaciones
            
    End Select
End Sub

Private Sub M0401040100_Click(Index As Integer)
    Select Case Index
        Case 0
            frmCredExtornos.Show 1
        Case 1
            frmCredExtPagoLote.Show 1
    End Select
End Sub

Private Sub M0601000000_Click(Index As Integer)
    Select Case Index
        Case 0 'Parametros
            
        Case 1 'Cosnsistema
             
        Case 2 'permisos
            frmMantPermisos.Show 1
        Case 3 'zonas
            frmMntZonas.Show 1
        Case 4 'Agencias
            frmMntAgencias.Show 1
        Case 5 'backUp
             
    End Select
End Sub

Private Sub M0701000000_Click(Index As Integer)
    If Index = 3 Then
        frmPosicionCli.Show 1
    End If
End Sub

Private Sub M0701010000_Click(Index As Integer)
    'Persona
    Select Case Index
        Case 0 'Registro
            frmPersona.Registrar
        Case 1 'mantenimiento
            frmPersona.Mantenimeinto
        Case 2 'Consulta
            frmPersona.Consultar
        Case 3 'Exoneradas del Lavado de Dinero
            frmPersLavDinero.Show 1
    End Select
End Sub

Private Sub M0701020000_Click(Index As Integer)
    
    'Instituciones Financieras
    Select Case Index
        Case 0
            frmMntInstFinanc.InicioActualizar
        Case 1
            frmMntInstFinanc.InicioConsulta
    End Select
End Sub

Private Sub M0701030000_Click(Index As Integer)

End Sub

Private Sub M0801000000_Click(Index As Integer)
    Select Case Index
        'Herramientas
        Case 1
            frmSpooler.Show 1
        Case 2
            frmSetupCOM.Show 1
    End Select
End Sub
