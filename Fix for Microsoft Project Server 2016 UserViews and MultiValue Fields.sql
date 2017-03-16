--
-- Fix for Microsoft Project Server 2016 UserViews and MultiValue Fields
--
-- [pjrep].[MSP_Epm_GenerateAllUserViews]
-- IF @SiteCount > 0  => IF @SiteCount > 1
--
-- [pjrep].[MSP_Epm_GenerateAllMultiValueAssociationViews]
-- CF_INDEX int IDENTITY (0,1), => CF_INDEX int IDENTITY (1,1),
--
-- USE AT YOUR OWN RISK - ABSOLUTLY NO WARRENTY
-- 
-- 
-- Run this on the Content-PWa-Database.
--
-- The script check for original content and replace it.
--
-- Hope that helps. 
-- Flori
--
-- f.grimm (at) solvin (dot) com
--
SET NOCOUNT ON;
BEGIN TRANSACTION
	DECLARE @forceUpdate BIT = 0;
	DECLARE @definitions TABLE(
		id INT IDENTITY(1,1),
		objectName nvarchar(max),
		definitionBuggy nvarchar(max),
		definitionNew nvarchar(max),
		PRIMARY KEY (id)
	)
	INSERT INTO @definitions (
	objectName,
	definitionBuggy,
	definitionNew
	) VALUES 

(
N'[pjrep].[MSP_Epm_GenerateAllUserViews]',
N'CREATE PROCEDURE [pjrep].[MSP_Epm_GenerateAllUserViews]
(
   @siteId uniqueidentifier
)
AS
BEGIN
   DECLARE @UserViewName   nvarchar(255)
   DECLARE @ResultValue    int   
   DECLARE @SiteCount      int
   EXEC pjrep.MSP_Epm_GetSiteCollectionCount @SiteCount OUT
   IF @SiteCount > 0
   BEGIN
      PRINT ''Skipping Reporting View creation. Reporting views are created for single tenant databases only''
      SET @ResultValue = 0
      RETURN @ResultValue
   END
   BEGIN TRANSACTION
   EXEC @ResultValue = pjrep.MSP_Epm_AquireSharedProjectMetadataLock @siteId
   IF @ResultValue != 0 GOTO LblError
   EXEC @ResultValue = pjrep.MSP_Epm_GenerateUserView @siteId, ''CECFE271-6660-4ABE-97ED-208D3C71FC18'', 0, @UserViewName OUT
   IF @ResultValue != 0 GOTO LblError
   EXEC @ResultValue = pjrep.MSP_Epm_GenerateUserView @siteId, ''EBAD93E7-2149-410D-9A39-A8680738329D'', 0, @UserViewName OUT
   IF @ResultValue != 0 GOTO LblError
   EXEC @ResultValue = pjrep.MSP_Epm_GenerateUserView @siteId, ''EBAD93E7-2149-410D-9A39-A8680738329D'', 1, @UserViewName OUT
   IF @ResultValue != 0 GOTO LblError
   EXEC @ResultValue = pjrep.MSP_Epm_GenerateUserView @siteId, ''C8C72436-F730-4443-B82B-52341ABFF84C'', 0, @UserViewName OUT
   IF @ResultValue != 0 GOTO LblError
   EXEC @ResultValue = pjrep.MSP_Epm_GenerateUserView @siteId, ''C8C72436-F730-4443-B82B-52341ABFF84C'', 1, @UserViewName OUT
   IF @ResultValue != 0 GOTO LblError
   EXEC @ResultValue = pjrep.MSP_Epm_GenerateUserView @siteId, ''BED1FDB9-6D08-4197-8BC2-9B6D6AA1D1CA'', 0, @UserViewName OUT
   IF @ResultValue != 0 GOTO LblError
   EXEC @ResultValue = pjrep.MSP_Epm_GenerateUserView @siteId, ''BED1FDB9-6D08-4197-8BC2-9B6D6AA1D1CA'', 1, @UserViewName OUT
   IF @ResultValue != 0 GOTO LblError
   COMMIT TRANSACTION
   RETURN 0
LblError:
   ROLLBACK TRANSACTION
   RETURN @ResultValue
END
',
N'CREATE PROCEDURE [pjrep].[MSP_Epm_GenerateAllUserViews]
(
   @siteId uniqueidentifier
)
AS
BEGIN
   DECLARE @UserViewName   nvarchar(255)
   DECLARE @ResultValue    int   
   DECLARE @SiteCount      int
   EXEC pjrep.MSP_Epm_GetSiteCollectionCount @SiteCount OUT
   IF @SiteCount > 1
   BEGIN
      PRINT ''Skipping Reporting View creation. Reporting views are created for single tenant databases only''
      SET @ResultValue = 0
      RETURN @ResultValue
   END
   BEGIN TRANSACTION
   EXEC @ResultValue = pjrep.MSP_Epm_AquireSharedProjectMetadataLock @siteId
   IF @ResultValue != 0 GOTO LblError
   EXEC @ResultValue = pjrep.MSP_Epm_GenerateUserView @siteId, ''CECFE271-6660-4ABE-97ED-208D3C71FC18'', 0, @UserViewName OUT
   IF @ResultValue != 0 GOTO LblError
   EXEC @ResultValue = pjrep.MSP_Epm_GenerateUserView @siteId, ''EBAD93E7-2149-410D-9A39-A8680738329D'', 0, @UserViewName OUT
   IF @ResultValue != 0 GOTO LblError
   EXEC @ResultValue = pjrep.MSP_Epm_GenerateUserView @siteId, ''EBAD93E7-2149-410D-9A39-A8680738329D'', 1, @UserViewName OUT
   IF @ResultValue != 0 GOTO LblError
   EXEC @ResultValue = pjrep.MSP_Epm_GenerateUserView @siteId, ''C8C72436-F730-4443-B82B-52341ABFF84C'', 0, @UserViewName OUT
   IF @ResultValue != 0 GOTO LblError
   EXEC @ResultValue = pjrep.MSP_Epm_GenerateUserView @siteId, ''C8C72436-F730-4443-B82B-52341ABFF84C'', 1, @UserViewName OUT
   IF @ResultValue != 0 GOTO LblError
   EXEC @ResultValue = pjrep.MSP_Epm_GenerateUserView @siteId, ''BED1FDB9-6D08-4197-8BC2-9B6D6AA1D1CA'', 0, @UserViewName OUT
   IF @ResultValue != 0 GOTO LblError
   EXEC @ResultValue = pjrep.MSP_Epm_GenerateUserView @siteId, ''BED1FDB9-6D08-4197-8BC2-9B6D6AA1D1CA'', 1, @UserViewName OUT
   IF @ResultValue != 0 GOTO LblError
   COMMIT TRANSACTION
   RETURN 0
LblError:
   ROLLBACK TRANSACTION
   RETURN @ResultValue
END
'
),
(
N'[pjrep].[MSP_Epm_GenerateAllMultiValueAssociationViews]',
N'CREATE PROCEDURE pjrep.MSP_Epm_GenerateAllMultiValueAssociationViews
   @siteId uniqueidentifier,
   @Result int OUTPUT
AS
BEGIN
   DECLARE @SiteCount           int
   DECLARE @Count               int
   DECLARE @i                   int
   DECLARE @CustomFieldTypeGuid uniqueidentifier
   EXEC pjrep.MSP_Epm_GetSiteCollectionCount @SiteCount OUT
   SET @Result = 0
   IF @SiteCount > 1
   BEGIN
      PRINT ''Skipping Reporting View creation. Reporting views are created for single tenant databases only''
      RETURN @Result 
   END
   CREATE TABLE #TBL_Reporting_MutliValuedCF(
      CF_INDEX int IDENTITY (0,1),
      CustomFieldTypeGuid uniqueidentifier)
   CREATE CLUSTERED INDEX CL ON #TBL_Reporting_MutliValuedCF(CustomFieldTypeGuid)
   SET @i = 1
   INSERT INTO #TBL_Reporting_MutliValuedCF(CustomFieldTypeGuid)
   SELECT AttributeTypeUID AS CustomFieldTypeGuid FROM pjrep.MSP_TVF_EpmMetadataAttribute(@siteId) 
   WHERE AttributeIsIntrinsic = 0 AND AttributeIsMultiValueEnabled = 1
   SELECT @Count = COUNT(CustomFieldTypeGuid) FROM #TBL_Reporting_MutliValuedCF 
   WHILE (@i <= @Count)
   BEGIN
   SELECT @CustomFieldTypeGuid = CustomFieldTypeGuid FROM #TBL_Reporting_MutliValuedCF WHERE CF_INDEX = @i
   EXEC @Result = pjrep.MSP_Epm_GenerateMultiValueAssociationView @siteId, @CustomFieldTypeGuid
   IF @Result != 0 GOTO LblError
   SELECT @i = @i + 1
   END
LblError:
   DROP TABLE #TBL_Reporting_MutliValuedCF
   RETURN @Result
END
',
N'CREATE PROCEDURE pjrep.MSP_Epm_GenerateAllMultiValueAssociationViews
   @siteId uniqueidentifier,
   @Result int OUTPUT
AS
BEGIN
   DECLARE @SiteCount           int
   DECLARE @Count               int
   DECLARE @i                   int
   DECLARE @CustomFieldTypeGuid uniqueidentifier
   EXEC pjrep.MSP_Epm_GetSiteCollectionCount @SiteCount OUT
   SET @Result = 0
   IF @SiteCount > 1
   BEGIN
      PRINT ''Skipping Reporting View creation. Reporting views are created for single tenant databases only''
      RETURN @Result 
   END
   CREATE TABLE #TBL_Reporting_MutliValuedCF(
      CF_INDEX int IDENTITY (1,1),
      CustomFieldTypeGuid uniqueidentifier)
   CREATE CLUSTERED INDEX CL ON #TBL_Reporting_MutliValuedCF(CustomFieldTypeGuid)
   SET @i = 1
   INSERT INTO #TBL_Reporting_MutliValuedCF(CustomFieldTypeGuid)
   SELECT AttributeTypeUID AS CustomFieldTypeGuid FROM pjrep.MSP_TVF_EpmMetadataAttribute(@siteId) 
   WHERE AttributeIsIntrinsic = 0 AND AttributeIsMultiValueEnabled = 1
   SELECT @Count = COUNT(CustomFieldTypeGuid) FROM #TBL_Reporting_MutliValuedCF 
   WHILE (@i <= @Count)
   BEGIN
   SELECT @CustomFieldTypeGuid = CustomFieldTypeGuid FROM #TBL_Reporting_MutliValuedCF WHERE CF_INDEX = @i
   EXEC @Result = pjrep.MSP_Epm_GenerateMultiValueAssociationView @siteId, @CustomFieldTypeGuid
   IF @Result != 0 GOTO LblError
   SELECT @i = @i + 1
   END
LblError:
   DROP TABLE #TBL_Reporting_MutliValuedCF
   RETURN @Result
END
'
);

	DECLARE @objectName nvarchar(max);
	DECLARE @definitionCurrent nvarchar(max);
	DECLARE @definitionBuggy nvarchar(max);
	DECLARE @definitionNew nvarchar(max);


	DECLARE cursorDefinitions CURSOR LOCAL READ_ONLY FOR SELECT objectName, definitionBuggy, definitionNew FROM @definitions;
	OPEN cursorDefinitions;

	FETCH NEXT FROM cursorDefinitions INTO @objectName,@definitionBuggy,@definitionNew
	WHILE (@@fetch_status <> -1)
	BEGIN
		IF (OBJECT_ID(@objectName) IS NULL) BEGIN
			PRINT ('ERROR   : ' + @objectName + ' not found.');
			GOTO LblError
		END ELSE BEGIN
			PRINT ('Info    : ' + @objectName + ' found.');
			SELECT @definitionCurrent=definition FROM [sys].[sql_modules] m WHERE m.object_id = OBJECT_ID(@objectName);
			IF ((@forceUpdate=1) OR (REPLACE(@definitionBuggy, ' ', '') = REPLACE(@definitionCurrent, ' ', ''))) BEGIN
				PRINT ('Info    : Orignal defintion of '+@objectName+' found.');
				SET @definitionNew = REPLACE(@definitionNew, 'CREATE PROCEDURE', 'ALTER PROCEDURE');
				BEGIN TRY
				EXECUTE sys.sp_executesql @definitionNew
				END TRY
				BEGIN CATCH
					GOTO LblError
				END CATCH

			END ELSE IF (REPLACE(@definitionNew, ' ', '') = REPLACE(@definitionCurrent, ' ', '')) BEGIN
				PRINT ('Info    : Modifed defintion of '+@objectName+' found.');
			END ELSE BEGIN
				PRINT ('Warning : Unknown Version of '+@objectName+' found')
				PRINT @definitionBuggy
				PRINT @definitionCurrent
			END;
		END;
		FETCH NEXT FROM cursorDefinitions INTO @objectName,@definitionBuggy,@definitionNew
	END

	CLOSE cursorDefinitions
	DEALLOCATE cursorDefinitions

	PRINT 'OK COMMIT TRANSACTION'
	COMMIT TRANSACTION
	GOTO LblEOF

LblError:
	PRINT 'ERROR ROLLBACK TRANSACTION'
	ROLLBACK TRANSACTION

LblEOF:
GO

/*

DON'T CALL
DECLARE @siteId uniqueidentifier;
DECLARE @ResultValue INT;
SELECT DISTINCT @siteId=SiteId FROM pjpub.MSP_WEB_ADMIN WHERE (1 = (SELECT Count = COUNT(*) FROM pjpub.MSP_WEB_ADMIN))
EXECUTE [pjrep].[MSP_Epm_RepairReportingViews] @siteId, @ResultValue OUTPUT
SELECT @ResultValue

INSTEAD CALL

PS> Repair-SPProjectWebInstance http://server/pwade -RepairRule ProjectRDBUserViewsRepairableHealthRule

*/