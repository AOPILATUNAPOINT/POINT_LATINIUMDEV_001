--exec sp_configure 'nested triggers', '0'
--reconfigure

-- Linea de comandos de actualiza estructura
--osql /U infoelect /P infoelect2243 /d BASEDEDATOS /S SERVIDOR\INFOELECT /n /i C:\Latinium\ActualizaVersion.Qry /o C:\Latinium\salida.txt

-- Proceso de Recuperar de suspect
-- checkpoint
-- EXEC sp_resetstatus 'bomberos'
-- DBCC CHECKDB ('bomberos', REPAIR_ALLOW_DATA_LOSS)
-- UPDATE master..sysdatabases SET status = status ^ 256 WHERE name = 'bomberos'
-- UPDATE master..sysdatabases SET status=-32768 WHERE name='bomberos'

-- Calculo de Total de Asiento
-- UPDATE Asiento SET Total = (SELECT Sum(Round(isnull(Debe, 0), 2)) 
-- FROM AsientoDetalle WHERE AsientoDetalle.idAsiento = Asiento.idAsiento)
-- UPDATE Asiento SET TotalH = (SELECT Sum(Round(isnull(Haber, 0), 2)) 
-- FROM AsientoDetalle WHERE AsientoDetalle.idAsiento = Asiento.idAsiento)


Set nocount on
Declare @BaseOrigen nVarChar(100), @BaseDestino nVarChar(100)
SET @BaseOrigen = 'USE Bases'
SET @BaseDestino = 'USE LatiniumSQL'
Declare @indice1 nvarchar(100), @campo1 nvarchar(100), @CreaIndice nvarchar(100)
Declare @idTabla1 int, @stTabla1 nchar(100)
Declare @stCol1 nchar(100), @Tipo1 int, @Long1 int, @Nulo1 int
Declare @idTabla2 int, @stTabla2 nchar(100)
Declare @stCol2 nchar(100), @Tipo2 int, @stTipo2 nVarchar(100), @Long2 int, @Nulo2 int
Declare @CreaTabla nVarChar(100), @Defa1 int, @Defa2 int, @MaxSeg int
-- Inserta Seguridad y Numeracion
SELECT @MaxSeg = Max(idTipoFactura) FROM LatiniumSQL.dbo.CompraNumero

Update LatiniumSQL.dbo.ProductoDetalle Set Proceso = 1 Where IsNull(Proceso, -1) = -1

insert into LatiniumSQL.dbo.CompraNumero 
	(idTipoFactura, Nombre, Numero, Codigo, Automatico, Abrev, Signo, Grupo, CobroPago)
	Select idTipoFactura,Nombre, Numero, Codigo, Automatico, Abrev, Signo, Grupo, CobroPago
	FROM bases.dbo.CompraNumero where idTipoFactura > @MaxSeg

SELECT @MaxSeg = Max(id_Seg) FROM LatiniumSQL.dbo.seguridad
insert into LatiniumSQL.dbo.seguridad
Select * FROM bases.dbo.seguridad where id_Seg > @MaxSeg

SELECT @MaxSeg = Max(idAuditaTipo) FROM LatiniumSQL.dbo.AuditaTipo

insert into LatiniumSQL.dbo.AuditaTipo (Tipo, Grupo)
	Select Tipo, grupo
	FROM bases.dbo.AuditaTipo where idAuditaTipo > @MaxSeg
-- Fin seguridad y numeracion

SELECT @MaxSeg = Max(idAsignacion) FROM LatiniumSQL.dbo.AsignaCuenta
insert into LatiniumSQL.dbo.AsignaCuenta (idAsignacion, Asignacion, idCuenta)
Select idAsignacion, Asignacion, idCuenta FROM bases.dbo.AsignaCuenta where idAsignacion > @MaxSeg

-- Seguridad
delete LatiniumSQL.dbo.Seguridad where id_seg in (373, 131, 244, 267)
delete LatiniumSQL.dbo.Seguridad where padre = 244
Update LatiniumSQL.dbo.Seguridad Set Codigo = '042RP', Padre = 626 where id_seg = 304
Update LatiniumSQL.dbo.Seguridad Set Codigo = '042PC', Padre = 626 where id_seg = 348
Update LatiniumSQL.dbo.Seguridad Set Codigo = '042PP', Padre = 626, Nombre = 'Pedido Proveedor' where id_seg = 310
Update LatiniumSQL.dbo.Seguridad Set Codigo = '042RC', Padre = 626, Nombre = 'Reservacion Cliente' where id_seg = 295
Update LatiniumSQL.dbo.Seguridad Set Nombre = 'Tarjeta credito' where id_seg = 580
Update LatiniumSQL.dbo.Seguridad Set Nombre = 'CxC y Pagar 30 60 90' where id_seg = 521
Update LatiniumSQL.dbo.Seguridad Set Nombre = 'Series' where id_seg = 574
Update LatiniumSQL.dbo.Seguridad Set Padre = 245 where id_seg = 268
-- Fin de seguridad

Update LatiniumSQL.dbo.CompraNumero Set Nombre = 'Reserva Cliente' where idTipoFactura = 11
Update LatiniumSQL.dbo.CompraNumero Set Nombre = 'Reserva Proveedor' where idTipoFactura = 26

-- Copia de Vistas
DECLARE curVistas CURSOR KEYSET FOR 
	SELECT v.TABLE_NAME, v.VIEW_DEFINITION from Information_Schema.VIEWS v Where Table_Name Like 'E_%'

DECLARE @Vista varchar(50), @Definicion varchar(4000)

OPEN curVistas
FETCH NEXT FROM curVistas INTO @Vista, @Definicion
WHILE (@@fetch_status <> -1)
BEGIN
	IF (@@fetch_status <> -2)
	BEGIN
		If Not Exists(SELECT * from Information_Schema.VIEWS v Where Table_Name = @Vista)
		Begin
			Exec (@Definicion)
		End
	END
	FETCH NEXT FROM curVistas INTO @Vista, @Definicion
END

CLOSE curVistas
DEALLOCATE curVistas
-- Fin de copia de vistas


DECLARE CurBases CURSOR READ_ONLY FOR 
	Select name, id from Bases.dbo.sysobjects where xtype='U' order by 1 

OPEN CurBases
FETCH NEXT FROM CurBases INTO @stTabla1, @idTabla1
WHILE (@@fetch_status <> -1)
BEGIN
  IF (@@fetch_status <> -2)
  BEGIN
-- Busqueda de Tabla Nueva
	IF NOT EXISTS (Select name from sysobjects where xtype='U' 
		And name=@stTabla1)
	BEGIN
	    SET @CreaTabla = 'SELECT * INTO ' + RTRIM(@stTabla1) + 
			' FROM Bases.dbo.' + RTRIM(@stTabla1)
	    EXEC (@CreaTabla)
	END
	ELSE
	BEGIN
		Select @idTabla2=id from sysobjects where xtype='U' 
			And name=@stTabla1
	END
-- Fin de Busqueda de Tabla Nueva
	
	DECLARE CurColumna1 CURSOR READ_ONLY For 
		Select name, xtype, length, isnullable, cDefault
			From Bases.dbo.SysColumns where id = @idTabla1

	OPEN CurColumna1
	FETCH NEXT FROM CurColumna1 
		INTO @stCol1, @Tipo1, @Long1, @Nulo1, @Defa1
	WHILE (@@fetch_status <> -1)
	BEGIN
	  IF (@@fetch_status <> -2)
	  BEGIN
		Declare @stDefa1 nvarchar(30), @stDefa2 nvarchar(30)
		SET @StDefa1 = ''
		SET @StDefa2 = ''
		select @stDefa1= text from Bases.dbo.syscomments 
	    	where id=@Defa1

		Declare @ModificaColumna int
		SET @ModificaColumna = 0
		SET @Defa2 = 0
		IF NOT EXISTS (Select Name
			From SysColumns where id = @idTabla2
			And Name = @stCol1)
		BEGIN
			SET @ModificaColumna = 1
		END
		ELSE -- Si existe verifica si se ha modificado
		BEGIN
		    Select @Tipo2 = SysColumns.xType, 
				@stTipo2 = SysTypes.Name, 
				@Long2 = SysColumns.length, 
				@Nulo2 = SysColumns.isnullable,
				@Defa2 = SysColumns.cDefault
		    	From SysColumns, SysTypes
				where id = @idTabla2
				And SysColumns.Name = @stCol1

		    select @stDefa2= text from syscomments 
			    where id=@Defa2

		    IF (@Nulo1 <> @Nulo2 Or @Tipo1 <> @Tipo2 Or @Long1 <> @Long2)
		    BEGIN
				IF @Long1 < @Long2
				BEGIN
					print 'Cambio en ' + RTRIM(@stTabla1) + ' ' + rtrim(@stCol1) 
						+ '- ' + @stDefa1 + ' - ' + @stDefa2 + ' Ancho Antes: ' 
						+ CAST(@Long2 as nvarChar(10)) + ' actual ' + CAST(@Long1 as nvarChar(10))
				END
				ELSE
				BEGIN
					print 'Cambio en ' + RTRIM(@stTabla1) + ' ' + rtrim(@stCol1) 
						+ '- ' + @stDefa1 + ' - ' + @stDefa2
					SET @ModificaColumna = 2
				END
		    END
		    ELSE
				IF (substring(@stDefa2, 1, 2) = '((')
					set @stDefa2 = substring(@stDefa2, 2, len(rtrim(@stDefa2))-2)

			IF (@stDefa2 <> @stDefa1)
				BEGIN
					If LEN(@stDefa2) > 0
					Begin
						DECLARE @defname VARCHAR(100)
						SET @defname =	(SELECT name
							FROM sysobjects so JOIN sysconstraints sc
							ON so.id = sc.constid
							WHERE object_name(so.parent_obj) = @stTabla1
							AND so.xtype = 'D'
							AND sc.colid =
							(SELECT colid FROM syscolumns
							WHERE id = object_id(@stTabla1) AND
							name = @stCol1))

						Set @CreaTabla = 'Alter Table ' + RTRIM(@stTabla1) 
							+ ' Drop Constraint ' + @defname
						Print @CreaTabla
						Exec (@CreaTabla)
					End

					IF LEN(RTRIM(IsNull(@stDefa1, ''))) > 0
					Begin
						Set @CreaTabla = 'Alter Table ' + RTRIM(@stTabla1) 
							+ ' Add Default ' +  RTRIM(@stDefa1) + ' For ' + rtrim(@stCol1) 
						Print @CreaTabla
						Exec (@CreaTabla)
					End
					SET @ModificaColumna = 3
			END
		END


		IF @ModificaColumna > 0 And @ModificaColumna < 3 -- Necesita modificar o crear una columna
		BEGIN -- Procedimiento Para agregar columnas
			Select @stTipo2 = SysTypes.Name From SysTypes where SysTypes.xtype = @Tipo1
			IF @stTipo2 = 'sysname'	SET @stTipo2 = 'nvarchar'
		  IF LEFT(@stTipo2,1) = 'n' SET @Long1 = @Long1/2
		  SET @CreaTabla = 'ALTER TABLE ' + RTRIM(@stTabla1) 
		  IF @ModificaColumna = 1
				SET @CreaTabla = @CreaTabla + ' ADD ' + RTRIM(@stCol1) + ' ' + @stTipo2
			ELSE
				SET @CreaTabla = @CreaTabla + ' ALTER COLUMN ' + RTRIM(@stCol1) + ' ' + @stTipo2

			If @stTipo2='varchar' Or @stTipo2='char'
				Or @stTipo2='nchar' Or @stTipo2='nvarchar'
			SET @CreaTabla = @CreaTabla + '(' + CAST(@Long1 as varchar(10)) +')'

			If @stTipo2='float' 
			SET @CreaTabla = @CreaTabla + '(53)'

			If @Nulo1 = 0
			SET @CreaTabla = @CreaTabla + ' NOT'
			SET @CreaTabla = @CreaTabla + ' NULL'
			Print @CreaTabla
			EXEC (@CreaTabla)
---------- Fin de Geracion de Columna

		-- Generacion de valores por defecto constantes como (0) o ((1.23)
		if LEN(@stDefa1) > 0 AND LEN(@stDefa1) <= 20
		BEGIN
		    SET @CreaTabla = 'ALTER TABLE ' + RTRIM(@stTabla1)
			+ ' ADD DEFAULT ' +  @stDefa1 + ' FOR ' + RTRIM(@stCol1)

		    print @CreaTabla
		    EXEC (@CreaTabla)
		END
		-- Asigna valor por defecto
		if LEN(@stDefa1) > 20
		BEGIN
--			print @stDefa1
		    SET @CreaTabla = 'EXEC sp_bindefault ' + char(39) +
			'Sucursal' + char(39) + ', ' + char(39) + 
			RTRIM(@stTabla1) + '.' + RTRIM(@stCol1) + char(39)
		    print @CreaTabla
		    EXEC (@CreaTabla)

		--- Modifica la tabla a no null
		    SET @CreaTabla = 'ALTER TABLE ' + RTRIM(@stTabla1) + 
			' ALTER COLUMN ' + RTRIM(@stCol1) + ' ' + 
			@stTipo2 
		    if @Nulo1 = 0
			SET @CreaTabla = @CreaTabla + ' NOT'
		    SET @CreaTabla = @CreaTabla + ' NULL'

		    print @CreaTabla
		    EXEC (@CreaTabla)
		END

		END
		FETCH NEXT FROM CurColumna1 INTO @stCol1, @Tipo1, @Long1, @Nulo1, @Defa1
	  END -- Fin de revision de Columnas
	end
	  CLOSE CurColumna1
	  DEALLOCATE CurColumna1
  -- Cierre de columnas


-- Revision de Indices
	DECLARE CurIndice1 CURSOR READ_ONLY For Select sk.keyno, si.name, sc.name 
		from Bases.dbo.sysindexes si, Bases.dbo.sysindexkeys sk, Bases.dbo.syscolumns sc
		where si.id=@idTabla1 
		And sk.id=@idTabla1
		And sk.indid = si.indid
		And sk.colid = sc.colid
		And sc.id = @idTabla1
	OPEN CurIndice1
	FETCH NEXT FROM CurIndice1 INTO @Long1, @Indice1, @Campo1 
	WHILE (@@fetch_status <> -1)
	BEGIN
	  IF (@@fetch_status <> -2)
	  BEGIN
		IF NOT EXISTS (Select Name
			From SysIndexes Where id = @idTabla2
			And Name = @Indice1)
		BEGIN -- Procedimiento Para agregar Indices
		    IF @Long1 = 1
			SET @CreaIndice = 'CREATE INDEX ' + @Indice1 
			+ ' ON ' + RTrim(@stTabla1) + ' (' + @Campo1 
		    ELSE
			SET @CreaIndice = @CreaIndice + ',' + @Campo1
		    SET @CreaIndice = @CreaIndice + ')'

--		    Print 'idt1: ' + cast(@idTabla1 as nvarchar(20)) + ': '
--		    + ' idt2: ' + cast(@idTabla2 as nvarchar(20)) + ': ' + @CreaIndice
--		    EXEC (@CreaIndice)
		END
		ELSE -- Existe la columna
		BEGIN
		    Select @Tipo2 = SysColumns.xType, 
				@stTipo2 = SysTypes.Name, 
				@Long2 = SysColumns.length, 
				@Nulo2 = SysColumns.isnullable,
				@Defa2 = SysColumns.cDefault
		    	From SysColumns, SysTypes 
				where id = @idTabla1
		    		And SysTypes.xtype = @Tipo1			
		END
	  END -- Fin de revision de indices

	  FETCH NEXT FROM CurIndice1 
		INTO @Long1, @Indice1, @Campo1 
	END	
	CLOSE CurIndice1
	DEALLOCATE CurIndice1
  

  FETCH NEXT FROM CurBases INTO @stTabla1, @idTabla1
  end
END

CLOSE CurBases
DEALLOCATE CurBases
Declare @vSelect VarChar(200)
SEt @vSelect = 'Update LatiniumSQL.dbo.Personal Set Tipo = 2 '
	+ 'Where IsNull(Tipo, 0) = 0'
Exec (@vSelect)
SEt @vSelect = 'Update LatiniumSQL.dbo.Articulo Set dCompra = 0, dVenta = 0 '
	+ 'Where IsNull(dCompra, 0) = 0 And IsNull(dVenta, 0) = 0'
Exec (@vSelect)
SEt @vSelect = 'Update LatiniumSQL.dbo.Retencion Set Tarjeta = 0, NoAplica = 0 '
	+ 'Where IsNull(Tarjeta, 0) = 0 And IsNull(NoAplica, 0) = 0'
Exec (@vSelect)
SEt @vSelect = 'Exec LatiniumSQL.dbo.CuentaNumPuntos '
Exec (@vSelect)

SEt @vSelect = 'Update LatiniumSQL.dbo.Cliente Set DesconProv = 0, DesconCliente = 0 '
	+ 'Where IsNull(DesconProv, 0) = 0 And IsNull(DesconCliente, 0) = 0'
Exec (@vSelect)

Set @vSelect = 'Update LatiniumSQL.dbo.Conciliacion Set Anterior = 1 '
	+ 'Where IsNull(Anterior, 1) = 1 And Confirmado = 1 '
Exec (@vSelect)
Set @vSelect = 'Update LatiniumSQL.dbo.Conciliacion Set Anterior = 1 '
	+ 'Where IsNull(Anterior, 1) = 1 And FechaConfirmacion > FechaGeneracion  '
Exec (@vSelect)
Set @vSelect = 'Update LatiniumSQL.dbo.Conciliacion Set Anterior = 0 '
	+ 'Where IsNull(Anterior, 0) = 0 '
Exec (@vSelect)
Set @vSelect = 'Update LatiniumSQL.dbo.Compra Set Servicio = 0 '
	+ 'Where IsNull(Servicio, -1) = -1 '
Exec (@vSelect)
Set @vSelect = 'Update LatiniumSQL.dbo.Compra Set Propina = 0 '
	+ 'Where IsNull(Propina, -1) = -1 '
Exec (@vSelect)
Declare @MinRolGrupo nVarChar(10)
Select @MinRolGrupo  = Cast(Min(idRolGrupo) as nvarchar(10)) From RolGrupo
Set @vSelect = 'Update LatiniumSQL.dbo.Personal Set idRolGrupo = ' + @MinRolGrupo 
	+ ' Where IsNull(idRolGrupo, 0) = 0 '
Exec (@vSelect)
GO


DECLARE @bErrors as bit
if not exists (select * from LatiniumSQL.dbo.sysobjects a inner join LatiniumSQL.dbo.sysindexes b
	on a.id = b.id  WHERE b.name = 'Auditoria2')
begin 
	BEGIN TRANSACTION
	SET @bErrors = 0

	CREATE NONCLUSTERED INDEX Auditoria2 ON LatiniumSQL.dbo.Auditoria (idUnica ASC, idTipoAudita ASC, 
		Fecha ASC, Campo ASC, Valor ASC, Estacion ASC, 			Usuario ASC, Codigo ASC )
	IF( @@error <> 0 ) SET @bErrors = 1

	IF( @bErrors = 0 )
	  COMMIT TRANSACTION
	ELSE
	  ROLLBACK TRANSACTION

end

Declare @MaxIdSeteoG int
SELECT @MaxIdSeteoG = Max(idSeteoG) FROM LatiniumSQL.dbo.SeteoG
Set Identity_Insert LatiniumSQL.dbo.SeteoG On 
Insert Into LatiniumSQL.dbo.SeteoG (idSeteoG, idPadre, Nombre, Activado, Valor, Codigo)
Select s.idSeteoG, idPadre, Nombre, Activado, Valor, Codigo FROM bases.dbo.SeteoG s where idSeteoG > @MaxIdSeteoG
Set Identity_Insert LatiniumSQL.dbo.SeteoG Off

Update Compra Set Fecha = Convert(DateTime, Convert(Char(10), Fecha,112), 112)
gO

ALTER TRIGGER [RETENCIONACTUALIZA] ON [dbo].[Retencion] 
FOR UPDATE, insert AS
------- Auditoria -----------
If Exists(SELECT * FROM SETEO Where AUDITORIA = 1)
Begin
	DECLARE @VIDUNICA INT, @VCOD VARCHAR(20), @VTIPO INT
	SELECT @VIDUNICA = i.IDRETENCION, @VCOD = i.CODIGO, @VTIPO = i.INGRESOEGRESO FROM INSERTED i
	SET @VTIPO = @VTIPO + 47

	Declare @VTABLA VARCHAR(100), @VCLAVE VARCHAR(100), @vEstado int, @vIdTipo int
	DECLARE @IDCOL BIGINT, @VCOL NVARCHAR(40), @VCOMPARA NVARCHAR(500), 
	@IDCOLN BIGINT, @IDPOS INT, @VSALIDA NVARCHAR(100)
	Set @VTABLA = 'RETENCION'
	Set @VCLAVE = 'IDRETENCION'
	Set @vEstado = 2
	Set @vIdTipo = @VTIPO
	SET @IDCOL = 1
	SET @IDPOS = 1
	SET @IDCOLN = 1

	SELECT @VCOL = COL_NAME(OBJECT_ID(@VTABLA), 1)
	WHILE @VCOL IS NOT NULL
	BEGIN
		IF (SUBSTRING(COLUMNS_UPDATED(), @IDPOS, 1) & @IDCOL) > 0 AND SUBSTRING(@VCOL,1,3) <>'UPS' --AND SUBSTRING(@VCOL,1,8) <>'FECHAEDI'
		BEGIN
		SET @VCOMPARA = 'SELECT @VSALIDA=' + @VCOL + 
		' FROM ' + @VTABLA + ' WHERE ' + @VCLAVE + ' = '
		+ CAST(@VIDUNICA AS VARCHAR(20))
		EXEC SP_EXECUTESQL  @VCOMPARA, N'@VSALIDA NVARCHAR(100) OUTPUT', @VSALIDA OUTPUT
		IF @VESTADO = 2 OR (@VSALIDA IS NOT NULL AND @VSALIDA <> '0')
			INSERT INTO AUDITORIA 
			(IDTIPOAUDITA, IDAUDITAESTADO, CODIGO, IDUNICA, CAMPO, VALOR)
			VALUES (@VIDTIPO, @VESTADO, @VCOD, @VIDUNICA, @VCOL, @VSALIDA)
		END
		SET @IDCOL = @IDCOL * 2
		SET @IDCOLN = @IDCOLN + 1
		IF @IDCOL >= 256     
		BEGIN
		SET @IDCOL = 1
		SET @IDPOS = @IDPOS + 1
		END
		SELECT @VCOL = COL_NAME(OBJECT_ID(@VTABLA), @IDCOLN)
	END
End
------- Fin de Auditoria -----------
UPDATE C SET C.RETENIDO = I.IVARETENIDO1 +
	I.IVARETENIDO + I.RENTARETENIDA1 + I.RENTARETENIDA
FROM COMPRA C INNER JOIN INSERTED I
ON C.IDCOMPRA = I.IDCOMPRA
