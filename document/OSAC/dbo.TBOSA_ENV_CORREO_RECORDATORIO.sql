USE [DBACOSAC_TEST]
GO
/****** Object:  StoredProcedure [dbo].[TBOSA_ENV_CORREO_RECORDATORIO]    Script Date: 21/02/2025 16:59:21 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- ================================================================
-- Author:    desarrollo1@acfarma.com
-- Create date: 2025-01-29
-- Description: Enviar correos electrónicos de recordatorio por cada Orden de Servicio en estado 4 (Pruebas)
-- ================================================================
CREATE PROCEDURE [dbo].[TBOSA_ENV_CORREO_RECORDATORIO]
AS
BEGIN

    DECLARE @id_orden NVARCHAR(150)
	DECLARE @codigo_orden NVARCHAR(50)
    DECLARE @id_solicitante NVARCHAR(150)
    DECLARE @solicitante NVARCHAR(150)
    DECLARE @correo NVARCHAR(150)
	DECLARE @tipoequipo NVARCHAR(150)
    DECLARE @equipo NVARCHAR(150)
	DECLARE @equipoResponsable NVARCHAR(255)
	DECLARE @area NVARCHAR(150)
    DECLARE @fecha_orden DATETIME
    DECLARE @fecha_programacion_inicio DATETIME
    DECLARE @fecha_programacion_fin DATETIME
    DECLARE @prioridad NVARCHAR(50)
    DECLARE @problema NVARCHAR(255)
    DECLARE @responsable NVARCHAR(150)
    DECLARE @fecha_inicio_atencion DATETIME
    DECLARE @fecha_fin_atencion DATETIME
    DECLARE @asunto NVARCHAR(200)
    DECLARE @body NVARCHAR(MAX)

    DECLARE correo_cursor CURSOR FOR
		SELECT DISTINCT 
			ROS.idOrdenServicio AS id, 
			ROS.codigo AS codigo, 
			ROS.idSolicitante AS idsolicitante, 
			ROS.solicitante AS solicitante, 
			USU.USUVCTXMAIL AS correo,
			ROS.tipoequipo AS tipoequipo, 
			ROS.nombre AS equipo, 
			(SELECT USS.USUVCNOUSUARIO AS responsable 
				FROM [DBACOSAC_TEST].[dbo].[TBOSA_EQUIPOS] AS EQU 
				INNER JOIN [DBACINHOUSE_TEST].[dbo].[TBSEGMAEUSUARIO] AS USS
				ON EQU.idUsuario = USS.USUINIDUSUARIO
				WHERE EQU.idEquipo = ORD.idEquipo) AS equipoResponsable,
			ARR.AREVCNOAREA AS area,
			ROS.fechaCreacion AS fechaCreacion,
			ROS.fechaProgramadaInicio AS fechaProgramadaInicio, 
			ROS.fechaProgramadaFin AS fechaProgramadaFin, 
			ROS.prioridad AS prioridad, 
			ROS.observacion AS problema,
			ROS.responsable AS responsable, 
			MAX(ROL.fechaInicio) AS fechaInicio,  
			MAX(ROL.fechaFin) AS fechaFin
		FROM 
			[DBACOSAC_TEST].[dbo].[VW_REPORTE_ORDENSERVICIO] AS ROS
		INNER JOIN
			[DBACOSAC_TEST].[dbo].[TBOSA_ORDENSERVICIO] AS ORD
			ON ROS.idOrdenServicio = ORD.idOrdenServicio
		INNER JOIN 
			[DBACINHOUSE_TEST].[dbo].[TBSEGMAEUSUARIO] AS USU
			ON ROS.idSolicitante = USU.USUINIDUSUARIO
		INNER JOIN [DBACINHOUSE_TEST].[dbo].[TBSEGMAEAREAS] AS ARR
			ON USU.AREINIDAREAPER = ARR.AREINIDAREA
		INNER JOIN 
			[DBACOSAC_TEST].[dbo].[VW_REPORTE_ORDENSERVICIO_ALL] ROL
			ON ROL.idOrdenServicio = ROS.idOrdenServicio
		WHERE 
			ROS.idEstado = '4' 
			AND USU.USUVCTXMAIL IS NOT NULL 
			AND USU.USUVCTXMAIL <> ''
		GROUP BY 
			ROS.idOrdenServicio, 
			ROS.codigo, 
			ROS.idSolicitante, 
			ROS.solicitante, 
			USU.USUVCTXMAIL,
			ROS.tipoequipo, 
			ROS.nombre, 
			ARR.AREVCNOAREA,
			ROS.fechaCreacion, 
			ROS.fechaProgramadaInicio, 
			ROS.fechaProgramadaFin, 
			ROS.prioridad, 
			ROS.observacion, 
			ROS.responsable, 
			ORD.idEquipo;

    OPEN correo_cursor

    FETCH NEXT FROM correo_cursor INTO 
        @id_orden,              
        @codigo_orden,              
        @id_solicitante,               
        @solicitante,               
        @correo,                    
        @tipoequipo,                
        @equipo, 
		@equipoResponsable,                   
        @area,                       
        @fecha_orden,               
        @fecha_programacion_inicio, 
        @fecha_programacion_fin,    
        @prioridad,                 
        @problema,                  
        @responsable,               
        @fecha_inicio_atencion,     
        @fecha_fin_atencion        

    WHILE @@FETCH_STATUS = 0
    BEGIN
        -- Sustituir valores NULL por valores por defecto o valores alternativos
        SET @problema = ISNULL(@problema, 'No disponible')
        SET @fecha_programacion_fin = ISNULL(@fecha_programacion_fin, @fecha_programacion_inicio)
        SET @fecha_fin_atencion = ISNULL(@fecha_fin_atencion, @fecha_inicio_atencion)

        -- Preparar el asunto y el cuerpo del correo para cada ticket
        SET @asunto = 'OSAC - Atención finalizada - ' + @codigo_orden
        SET @body = 
			'Estimado(a) <strong>' + @solicitante + '</strong>,<br /><br />' +
			'Nos complace informarle que se ha finalizado la atención de su Orden de Servicio. <br />' +
			'Este es un recordatorio para que por favor verifique si el problema ha sido completamente solucionado. <br />Una vez hecho esto, ingrese al sistema OSAC para indicarnos si está conforme con la solución o si persiste algún inconveniente. <br />Recuerde que usted puede realizar esta acción en el transcurso de 24 horas.<br /><br />' +
			'<strong>Tipo de orden:</strong> Orden de Servicio <br />' +
			'<strong>Código:</strong> ' + @codigo_orden + '<br />' +
			'<strong>Solicitante:</strong> ' + @solicitante + '<br />' +
			'<strong>Área:</strong> ' + @area + '<br />' +
			'<strong>Tipo de equipo:</strong> ' + @tipoequipo + '<br />' +
			'<strong>Equipo:</strong> ' + @equipo + '<br />' +
			'<strong>Asignado a:</strong> ' + @equipoResponsable + '<br />' +
			'<strong>Fecha de la orden:</strong> ' + 
			CONVERT(VARCHAR, @fecha_orden, 103) + ' ' + 
			RIGHT('0' + CAST(DATEPART(HOUR, @fecha_orden) % 12 AS VARCHAR), 2) + ':' + 
			RIGHT('0' + CAST(DATEPART(MINUTE, @fecha_orden) AS VARCHAR), 2) + ' ' + 
			CASE WHEN DATEPART(HOUR, @fecha_orden) < 12 THEN 'AM' ELSE 'PM' END + '<br />' +
			'<strong>Fecha de programación:</strong> ' + 
			CONVERT(VARCHAR, @fecha_programacion_inicio, 103) + ' ' + 
			RIGHT('0' + CAST(DATEPART(HOUR, @fecha_programacion_inicio) % 12 AS VARCHAR), 2) + ':' + 
			RIGHT('0' + CAST(DATEPART(MINUTE, @fecha_programacion_inicio) AS VARCHAR), 2) + ' ' + 
			CASE WHEN DATEPART(HOUR, @fecha_programacion_inicio) < 12 THEN 'AM' ELSE 'PM' END + 
			' hasta ' +
			CONVERT(VARCHAR, @fecha_programacion_fin, 103) + ' ' + 
			RIGHT('0' + CAST(DATEPART(HOUR, @fecha_programacion_fin) % 12 AS VARCHAR), 2) + ':' + 
			RIGHT('0' + CAST(DATEPART(MINUTE, @fecha_programacion_fin) AS VARCHAR), 2) + ' ' + 
			CASE WHEN DATEPART(HOUR, @fecha_programacion_fin) < 12 THEN 'AM' ELSE 'PM' END + '<br />' +
			'<strong>Prioridad:</strong> ' + @prioridad + '<br />' +
			'<strong>Problema:</strong> ' + @problema + '<br /><br />' +
			'<strong>Detalle de la atención:</strong><br />' +
			'<strong>Responsable:</strong> ' + @responsable + '<br />' +
			'<strong>Fecha inicio:</strong> ' + 
			CONVERT(VARCHAR, @fecha_inicio_atencion, 103) + ' ' + 
			RIGHT('0' + CAST(DATEPART(HOUR, @fecha_inicio_atencion) % 12 AS VARCHAR), 2) + ':' + 
			RIGHT('0' + CAST(DATEPART(MINUTE, @fecha_inicio_atencion) AS VARCHAR), 2) + ' ' + 
			CASE WHEN DATEPART(HOUR, @fecha_inicio_atencion) < 12 THEN 'AM' ELSE 'PM' END + '<br />' +
			'<strong>Fecha fin:</strong> ' + 
			CONVERT(VARCHAR, @fecha_fin_atencion, 103) + ' ' + 
			RIGHT('0' + CAST(DATEPART(HOUR, @fecha_fin_atencion) % 12 AS VARCHAR), 2) + ':' + 
			RIGHT('0' + CAST(DATEPART(MINUTE, @fecha_fin_atencion) AS VARCHAR), 2) + ' ' + 
			CASE WHEN DATEPART(HOUR, @fecha_fin_atencion) < 12 THEN 'AM' ELSE 'PM' END + '<br /><br />' +
			'Atte.<br />' +
			'Administrador de Software.<br />' +
			'OSAC - Sistema Órdenes de Servicio<br /><br /><br /><br />' +
			'Antes de imprimir este mensaje, piensa!, Realmente lo necesitas?'

        -- Enviar el correo para cada ticket
        IF @correo IS NOT NULL AND LTRIM(RTRIM(@correo)) <> '' AND @asunto IS NOT NULL AND LTRIM(RTRIM(@asunto)) <> '' AND @body IS NOT NULL AND LTRIM(RTRIM(@body)) <> ''
        BEGIN
            EXEC msdb.dbo.sp_send_dbmail
                @profile_name = 'Developer Profile',  -- Cambiar al nombre de tu perfil de correo
                @recipients = @correo,
                @subject = @asunto,
                @body = @body,
                @body_format = 'HTML',
                @exclude_query_output = 1
        END
        ELSE
        BEGIN
            PRINT 'Correo vacío o sin contenido. No se enviará correo a: ' + @correo
        END

        -- Recuperar el siguiente conjunto de resultados
        FETCH NEXT FROM correo_cursor INTO 
            @id_orden,              
            @codigo_orden,              
            @id_solicitante,               
            @solicitante,               
            @correo,                    
            @tipoequipo,                
            @equipo,  
			@equipoResponsable,                   
            @area,                       
            @fecha_orden,               
            @fecha_programacion_inicio, 
            @fecha_programacion_fin,    
            @prioridad,                 
            @problema,                  
            @responsable,               
            @fecha_inicio_atencion,     
            @fecha_fin_atencion        
    END

    -- Cerrar el cursor
    CLOSE correo_cursor
    DEALLOCATE correo_cursor
END
