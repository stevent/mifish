CREATE DATABASE IF NOT EXISTS `simplistik`;
USE `simplistik`;
GRANT ALL PRIVILEGES  ON `simplistik` TO `simplistik`@`localhost` IDENTIFIED BY 'u78FiU70Pi' WITH GRANT OPTION;


DROP TABLE IF EXISTS `page`;
CREATE TABLE `page` (
	`Id` int(11) NOT NULL auto_increment,
	`ParentID` int(11) NOT NULL,
	`Name` varchar(50) NOT NULL,
  `Permalink` varchar(50) NOT NULL,
  `Heading` varchar(255) NULL,
	`PageTitle` varchar(255) NULL,
	`MetaDesc` varchar(255) NULL,
  `Active` tinyint(1) NOT NULL DEFAULT 1,
  `IsHomePage` tinyint(1) NOT NULL DEFAULT 0,
	`Rank` int(11) NOT NULL,
	`Content` text NULL,
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB AUTO_INCREMENT=24 DEFAULT CHARSET=latin1;
INSERT INTO `page` ( Id, ParentID, Name, Permalink, Heading, Active, IsHomePage, Rank, Content ) VALUES ( 1, 0, 'Home', '', 'Home', true, true, 1, '<p>Welcome to Simplistik Development' );



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[page]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[page];

CREATE TABLE [dbo].[page] (
	[Id] [int] IDENTITY (1, 1) NOT FOR REPLICATION  PRIMARY KEY,
  [ParentID] [int]  NOT NULL,
	[Name] [varchar] (50) COLLATE Latin1_General_CI_AS NOT NULL,
  [Permalink] [varchar] (50) COLLATE Latin1_General_CI_AS NULL,
  [Heading] [varchar] (250) COLLATE Latin1_General_CI_AS NOT NULL,
  [Active] [bit] NOT NULL,
  [IsHomePage] [bit] NOT NULL,
  [Rank] [int] NOT NULL,
	[Content] [text] COLLATE Latin1_General_CI_AS NULL,
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY];
INSERT INTO page (ParentID, Name, Permalink, Heading, Active, IsHomePage, Rank, Content) VALUES (0, 'Home', '', 'Home', 1, 1, 1, '');
