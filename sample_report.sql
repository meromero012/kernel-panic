/*************************************************
 *************** STORED PROCEDURES ***************
 *************************************************/

CREATE OR ALTER PROCEDURE [dbo].[WriteLNToFile]
	@FilePath	VARCHAR[4000],
	@Text		VARCHAR[4000]
AS
BEGIN
	DECLARE @OLE	INT
	DECLARE @FileID	INT
	
	EXEC sp_OACreate 'Scripting.FileSystemObject', @OLE OUT
	EXEC sp_OAMethod @OLE, 'OpenTextFile', @FileID OUT, @File, 8, 1
	EXEC sp_OAMethod @FileID, 'WriteLine', NULL, @TEXT
	
	EXEC sp_OADestroy @FileID
	EXEC sp_OADestroy @OLE
END
GO;

CREATE OR ALTER PROCEDURE [dbo].[SetOleAutoProcOn]
	@ValueInUse INT
AS
BEGIN
	IF NOT(@ValueInUse = 1)
	BEGIN
		EXEC sp_configure 'show advanced options', 1
		RECONFIGURE
		EXEC sp_configure 'Ole Automation Procedures', 1
		RECONFIGURE
	END
END
GO;

CREATE OR ALTER PROCEDURE [dbo].[SetOleAutoProcOff]
	@PreviousValue INT
AS
BEGIN
	DECLARE @ValueInUse INT

	SELECT @ValueInUse = [VALUE_IN_USE]
	FROM [SYS].[CONFIGURATIONS]
	WHERE [NAME] = 'Ole Automation Procedures'
	
	IF NOT(@ValueInUse <> @PreviousValue)
	BEGIN
		EXEC sp_configure 'show advanced options', 1
		RECONFIGURE
		EXEC sp_configure 'Ole Automation Procedures', @PreviousValue
		RECONFIGURE
	END
END
GO;

CREATE OR ALTER PROCEDURE [dbo].[SampleReport]
	@FilePath 	VARCHAR[255]
AS
BEGIN
	-- VARIABLE DECLARATIONS

	DECLARE @ValueInUse INT

	-- VARIABLE INITIALIZATIONS

	SET @FilePath = 'C:/SCRIPT/Script_Result.html'

	SELECT @ValueInUse = [VALUE_IN_USE]
	FROM [SYS].[CONFIGURATIONS]
	WHERE [NAME] = 'Ole Automation Procedures'

	-- SET OLE AUTOMATION PROCEDURES ON

	EXEC [dbo].[SetOleAutoProcOn] @ValueInUse

	-- WRITE HTML HEADER

	EXEC [dbo].[WriteLNToFile] @FilePath, '<!DOCTYPE html>'
	EXEC [dbo].[WriteLNToFile] @FilePath, '<html>'
	EXEC [dbo].[WriteLNToFile] @FilePath, '<head>'
	EXEC [dbo].[WriteLNToFile] @FilePath, '<title> SAMPLE SCRIPT </title>'
	EXEC [dbo].[WriteLNToFile] @FilePath, '<style> table, th, td {border: 1px solid black; border-collapse: collapse;}'
	EXEC [dbo].[WriteLNToFile] @FilePath, 'th, td {padding: 5px;} </style>'
	EXEC [dbo].[WriteLNToFile] @FilePath, '</head>'
	EXEC [dbo].[WriteLNToFile] @FilePath, '<body>'
	EXEC [dbo].[WriteLNToFile] @FilePath, '<h1> SAMPLE SCRIPT </h1>'
	EXEC [dbo].[WriteLNToFile] @FilePath, '<p> Execution Timestamp: ' + CONVERT(VARCHAR[19], GETDATE(), 114) + '</p>'

	-- TEST TABLE

	EXEC [dbo].[WriteLNToFile] @FilePath, '<h2> Table: SYS.CONFIGURATIONS </h2>'
	EXEC [dbo].[WriteLNToFile] @FilePath, '<p> Executed Query: </p>'
	EXEC [dbo].[WriteLNToFile] @FilePath, '<p> SELECT NAME, VALUE, RUN_VALUE FROM SYS.CONFIGURATIONS WHERE NAME = ''Ole Automation Procedures'' </p>'
	EXEC [dbo].[WriteLNToFile] @FilePath, '<table style="width:100%">'
	EXEC [dbo].[WriteLNToFile] @FilePath, '<tr> <th>NAME</th> <th>VALUE</th> <th>RUN_VALUE</th> </tr>'

	DECLARE @Name		NVARCHAR[35],
			@Value		NVARCHAR[MAX],
			@RunValue	NVARCHAR[MAX]

	DECLARE C_SysConfigs CURSOR FOR
	SELECT [NAME], [VALUE], [RUN_VALUE]
	FROM [SYS].[CONFIGURATIONS]
	WHERE [NAME] = 'Ole Automation Procedures';

	OPEN C_SysConfigs
	FETCH NEXT FROM C_SysConfigs
	INTO @Name, @Value, @RunValue

	WHILE @@FETCH_STATUS = 0
	BEGIN

		EXEC [dbo].[WriteLNToFile] @FilePath, '<tr> <td>' + @Name + '</td> <td>' + @Value + '</td> <td>' + @RunValue + '</td> </tr>'

		FETCH NEXT FROM C_SysConfigs
		INTO @Name, @Value, @RunValue
	END

	CLOSE C_SysConfigs;
	DEALLOCATE C_SysConfigs;

	EXEC [dbo].[WriteLNToFile] @FilePath, '</table>'

	-- WRITE HTML ENDING

	EXEC [dbo].[WriteLNToFile] @FilePath, '</body>'
	EXEC [dbo].[WriteLNToFile] @FilePath, '</html>'

	-- SET OLE AUTOMATION PROCEDURES OFF

	EXEC [dbo].[SetOleAutoProcOff] @ValueInUse
END
GO;
