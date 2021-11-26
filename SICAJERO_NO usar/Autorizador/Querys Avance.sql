USE [DBTarjeta]
GO
/****** Objeto:  Table [dbo].[ATM_Sucesos]    Fecha de la secuencia de comandos: 10/06/2007 20:16:56 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[ATM_Sucesos](
	[dFecha] [datetime] NOT NULL,
	[cProceso] [varchar](150) COLLATE Modern_Spanish_CI_AS NOT NULL,
	[cDescripcion] [varchar](8000) COLLATE Modern_Spanish_CI_AS NOT NULL
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF

USE [DBTarjeta]
GO
/****** Objeto:  Table [dbo].[Tarjeta]    Fecha de la secuencia de comandos: 10/06/2007 20:17:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Tarjeta](
	[cNumTarjeta] [varchar](20) COLLATE Modern_Spanish_CI_AS NOT NULL,
	[nCondicion] [int] NOT NULL,
	[nRetenerTarjeta] [int] NOT NULL,
 CONSTRAINT [PK_Tarjeta] PRIMARY KEY CLUSTERED 
(
	[cNumTarjeta] ASC
)WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF

USE [DBTarjeta]
GO
/****** Objeto:  Table [dbo].[Tarjeta_Param]    Fecha de la secuencia de comandos: 10/06/2007 20:17:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Tarjeta_Param](
	[nNOOperMonExt] [smallint] NOT NULL,
	[nSuspOper] [smallint] NOT NULL CONSTRAINT [DF_Tarjeta_Param_nSuspOper]  DEFAULT ((0))
) ON [PRIMARY]

USE [DBTarjeta]
GO
/****** Objeto:  StoredProcedure [dbo].[ATM_RecuperaDatosCuenta]    Fecha de la secuencia de comandos: 10/06/2007 20:17:49 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [dbo].[ATM_RecuperaDatosCuenta]
@psCtaCod char(18), @nSaldo money OUTPUT
AS
BEGIN
Declare @lnSaldoMinimoCta as money
Declare @lnSaldoDisp as money
Declare @lnRetencion as money
declare @lnPersoneria as int
declare @lbOrdenPago as bit

	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;
	
	
/*	Select @lnSaldo = PRD.nSaldo, @lnEstado = PRD.nPRDEstado, @lnSaldoDisp = CAP.nSaldoDisp, @lnRetencion = ISNULL(AHO.nRetencion,0), 
		   @ldUltCierre = CAP.dUltCierre, @lnTransacc = PRD.nTransacc, @lnTasaInteres = PRD.nTasaInteres, 
		   @lnIntAcum = CAP.nIntAcum, @lbInactiva = AHO.bInactiva, @lbInmovilizada  = AHO.bInmovilizada, @lnPersoneria = CAP.nPersoneria, @lbOrdenPago = AHO.bOrdPag
*/
	Select @lnSaldoDisp = CAP.nSaldoDisp, @lnRetencion = ISNULL(AHO.nRetencion,0), 
		@lnPersoneria = CAP.nPersoneria, @lbOrdenPago = AHO.bOrdPag
	From DBNegocio..Producto PRD
	Inner Join DBNegocio..Captaciones CAP On PRD.cCtaCod = CAP.cCtaCod
	Inner Join DBNegocio..CaptacAhorros AHO On AHO.cCtaCod = CAP.cCtaCod
	Where PRD.cCtaCod = @psCtaCod And PRD.nPRDEstado = 1000

	Select @lnSaldoMinimoCta = nSaldoMin 
	From DBNegocio..CapPersParam 
	Where nPersoneria = @lnPersoneria And nProducto = SubString(@psCtaCod,6,3) 
	And nMoneda = SubString(@psCtaCod,9,1) 
	And cOrdPag = @lbOrdenPago

	If @lnSaldoMinimoCta Is Null 
		Set @lnSaldoMinimoCta = 0

	SET @nSaldo = @lnSaldoMinimoCta + @lnSaldoDisp - @lnRetencion 

END

USE [DBTarjeta]
GO
/****** Objeto:  StoredProcedure [dbo].[ATM_RecuperaDatosTarjeta]    Fecha de la secuencia de comandos: 10/06/2007 20:17:57 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO




-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [dbo].[ATM_RecuperaDatosTarjeta] 
@PAN varchar(20), @nCondicion int OUTPUT, @nRetenerTarjeta int OUTPUT, 
@nNOOperMonExt int OUTPUT, @nSuspOper int OUTPUT
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

	Select @nCondicion=nCondicion, @nRetenerTarjeta=nRetenerTarjeta 
	From tarjeta Where cNumTarjeta=@PAN
	Set @nCondicion = isnull(@nCondicion,0)	

	Select @nNOOperMonExt=nNOOperMonExt,@nSuspOper=nSuspOper 
	from Tarjeta_Param
	
END

USE [DBTarjeta]
GO
/****** Objeto:  StoredProcedure [dbo].[ATM_RegistraSucesos]    Fecha de la secuencia de comandos: 10/06/2007 20:18:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [dbo].[ATM_RegistraSucesos]
@dFecha datetime, @cProceso varchar(150), @cDescripcion varchar(5000)
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

	INSERT INTO ATM_Sucesos(dFecha, cProceso,cDescripcion)
	VALUES(@dFecha,@cProceso,@cDescripcion)

END


