/******************************

sql-scripting 08/20/2001 

email: sqlscripters@sql-scripting.com  

Visit www.sql-scripting.com 

******************************/ 

USE yfymis

GO

SET QUOTED_IDENTIFIER OFF SET ANSI_NULLS ON 

GO 

-- CREATE PROCEDURE sp_DBA_DBCCShowFragAll  AS 

BEGIN 

DECLARE UserTables INSENSITIVE CURSOR 

     FOR 

  Select    name FROM sysobjects --select table names

  WHERE type = 'U'  and left(name,1)='y'  and left(name,3)<>'ycb' 

  ORDER BY name 

FOR READ ONLY 

OPEN UserTables 

DECLARE @TableName varchar(50), 

                  @MSG varchar(max), 

                  @id int 

FETCH NEXT FROM UserTables INTO @TableName --pass tbl names

   WHILE (@@FETCH_STATUS = 0)--loop through tablenames 

      BEGIN 

SELECT @MSG = 'DBCC SHOWCONTIG For table: ' + @TableName 

 PRINT @MSG --print some info

SET @id = object_id(@tablename)--set variable to pass 

DBCC SHOWCONTIG (@id) --execute
print '------------------------------------------------------------------------------------'
  

FETCH NEXT FROM UserTables INTO @TableName 

END 

CLOSE UserTables 

DEALLOCATE UserTables 

END 

GO 

SET QUOTED_IDENTIFIER OFF SET ANSI_NULLS ON 

GO 

 