IF EXISTS (SELECT name FROM master.dbo.sysdatabases WHERE name = N'Notare')
	DROP DATABASE [Notare]
GO

CREATE DATABASE [Notare]
 COLLATE Latin1_General_CI_AS
GO

exec sp_dboption N'Notare', N'autoclose', N'false'
GO

exec sp_dboption N'Notare', N'bulkcopy', N'false'
GO

exec sp_dboption N'Notare', N'trunc. log', N'false'
GO

exec sp_dboption N'Notare', N'torn page detection', N'true'
GO

exec sp_dboption N'Notare', N'read only', N'false'
GO

exec sp_dboption N'Notare', N'dbo use', N'false'
GO

exec sp_dboption N'Notare', N'single', N'false'
GO

exec sp_dboption N'Notare', N'autoshrink', N'false'
GO

exec sp_dboption N'Notare', N'ANSI null default', N'false'
GO

exec sp_dboption N'Notare', N'recursive triggers', N'false'
GO

exec sp_dboption N'Notare', N'ANSI nulls', N'false'
GO

exec sp_dboption N'Notare', N'concat null yields null', N'false'
GO

exec sp_dboption N'Notare', N'cursor close on commit', N'false'
GO

exec sp_dboption N'Notare', N'default to local cursor', N'false'
GO

exec sp_dboption N'Notare', N'quoted identifier', N'false'
GO

exec sp_dboption N'Notare', N'ANSI warnings', N'false'
GO

exec sp_dboption N'Notare', N'auto create statistics', N'true'
GO

exec sp_dboption N'Notare', N'auto update statistics', N'true'
GO

if( (@@microsoftversion / power(2, 24) = 8) and (@@microsoftversion & 0xffff >= 724) )
	exec sp_dboption N'Notare', N'db chaining', N'false'
GO

use [Notare]
GO

