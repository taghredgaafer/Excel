SET ANSI_NULLS ON

GO

SET QUOTED_IDENTIFIER ON

GO

CREATE PROCEDURE spx_ImportFromExcel16

   @SheetName varchar(20),

   @FilePath varchar(100),

   @HDR varchar(3),

   @TableName varchar(50)

AS

BEGIN

    DECLARE @SQL nvarchar(1000)

     

    IF OBJECT_ID (@TableName,'U') IS NOT NULL

      SET @SQL = 'INSERT INTO ' + @TableName + ' SELECT * FROM OPENDATASOURCE'

    ELSE

      SET @SQL = 'SELECT * INTO ' + @TableName + ' FROM OPENDATASOURCE'

 

    SET @SQL = @SQL + '(''Microsoft.ACE.OLEDB.16.0'',''Data Source='

    SET @SQL = @SQL + @FilePath + ';Extended Properties=''''Excel 16.0;HDR='

    SET @SQL = @SQL + @HDR + ''''''')...['

    SET @SQL = @SQL + @SheetName + ']'

    EXEC sp_executesql @SQL

END

GO
------------------------------

EXEC master.dbo.sp_MSset_oledb_prop N'Microsoft.ACE.OLEDB.12.0'
 
    , N'AllowInProcess', 1
 
GO


EXEC master.dbo.sp_MSset_oledb_prop N'Microsoft.ACE.OLEDB.12.0'

    , N'DynamicParameters', 1

GO