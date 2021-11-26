DECLARE @sTipo VarChar(1)
DECLARE @sCta VarChar(12)

DECLARE Cuenta_Cursor CURSOR FOR
Select MAX(Tipo) Pers, Cuenta FROM (
Select P.cTipPers Tipo, pc.cCodCta Cuenta FROM AhorroC A INNER JOIN PersCuenta PC 
INNER JOIN DBPersona..Persona P ON PC.cCodPers = P.cCodPers ON A.cCodCta = PC.cCodCta
WHERE (A.cPersoneria = '' OR A.cPersoneria = '0') AND PC.cRelaCta = 'TI'
UNION
Select P.cTipPers Tipo, pc.cCodCta Cuenta FROM PlazoFijo A INNER JOIN PersCuenta PC 
INNER JOIN DBPersona..Persona P ON PC.cCodPers = P.cCodPers ON A.cCodCta = PC.cCodCta
WHERE (A.cPersoneria = '' OR A.cPersoneria = '0') AND PC.cRelaCta = 'TI'
UNION
Select P.cTipPers Tipo, pc.cCodCta Cuenta FROM CTS A INNER JOIN PersCuenta PC 
INNER JOIN DBPersona..Persona P ON PC.cCodPers = P.cCodPers ON A.cCodCta = PC.cCodCta
WHERE (A.cPersoneria = '' OR A.cPersoneria = '0') AND PC.cRelaCta = 'TI'
) T Group by Cuenta


SET NOCOUNT ON
OPEN Cuenta_Cursor

Begin Transaction
FETCH Cuenta_Cursor INTO  @sTipo, @sCta
WHILE @@FETCH_STATUS = 0
  BEGIN
	if (@sTipo = '4' Or @sTipo = '5') Select @sTipo = '3'
	If Substring(@sCta,3,3) = "232" Update AhorroC Set cPersoneria = @sTipo Where cCodCta = @sCta
	If Substring(@sCta,3,3) = "233" Update PlazoFijo Set cPersoneria = @sTipo Where cCodCta = @sCta
	If Substring(@sCta,3,3) = "234" Update CTS Set cPersoneria = @sTipo Where cCodCta = @sCta
	PRINT 'Cuenta : ' + @sCta + '   Tipo : ' + @sTipo
	FETCH Cuenta_Cursor INTO  @sTipo, @sCta
  END
Commit Transaction

CLOSE Cuenta_Cursor
DEALLOCATE Cuenta_Cursor
SET NOCOUNT OFF
PRINT 'Proceso Finalizado con éxito. Podeis ir en Paz'






