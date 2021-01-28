USE [Marketing]
GO

/****** Object:  Table [dbo].[Developers_Staging]    Script Date: 1/27/2021 4:41:04 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[Developers_Staging](
	[TitleID] [float] NULL,
	[TitleName] [nvarchar](255) NULL,
	[Developers] [nvarchar](255) NULL,
	[StudioID] [nvarchar](255) NULL
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[GamesTitles]    Script Date: 1/27/2021 4:41:04 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[GamesTitles](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[TitleID] [nvarchar](255) NULL,
	[TitleName] [nvarchar](255) NULL,
	[Metascore] [nvarchar](255) NULL,
	[GameModes] [nvarchar](255) NULL,
	[Genre] [nvarchar](255) NULL,
	[Themes] [nvarchar](255) NULL,
	[Series] [nvarchar](255) NULL,
	[PlayerPerspectives] [nvarchar](255) NULL,
	[Franchises] [nvarchar](255) NULL,
	[GameEngine] [nvarchar](255) NULL,
	[AlternativeNames] [nvarchar](255) NULL,
	[IGDB_Website] [nvarchar](255) NULL,
	[NewZoo_Website] [nvarchar](255) NULL
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[GamesTitles_Dev]    Script Date: 1/27/2021 4:41:04 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[GamesTitles_Dev](
	[UID] [int] IDENTITY(1,1) NOT NULL,
	[ID] [nvarchar](255) NOT NULL,
	[TitleID] [nvarchar](255) NULL,
	[TitleName] [nvarchar](255) NULL,
	[IGDB_Website] [nvarchar](255) NULL,
	[NewZoo_Website] [nvarchar](255) NULL,
	[Developer] [nvarchar](255) NULL,
	[StudioID] [nvarchar](255) NOT NULL,
	[VTSID] [nvarchar](255) NOT NULL
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[GamesTitles_PortDev]    Script Date: 1/27/2021 4:41:04 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[GamesTitles_PortDev](
	[UID] [int] IDENTITY(1,1) NOT NULL,
	[ID] [nvarchar](255) NOT NULL,
	[TitleID] [nvarchar](255) NULL,
	[TitleName] [nvarchar](255) NULL,
	[IGDB_Website] [nvarchar](255) NULL,
	[NewZoo_Website] [nvarchar](255) NULL,
	[PortDev] [nvarchar](255) NULL,
	[StudioID] [nvarchar](255) NOT NULL,
	[VTSID] [nvarchar](255) NOT NULL
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[GamesTitles_Pub]    Script Date: 1/27/2021 4:41:04 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[GamesTitles_Pub](
	[UID] [int] IDENTITY(1,1) NOT NULL,
	[ID] [nvarchar](255) NOT NULL,
	[TitleID] [nvarchar](255) NULL,
	[TitleName] [nvarchar](255) NULL,
	[IGDB_Website] [nvarchar](255) NULL,
	[NewZoo_Website] [nvarchar](255) NULL,
	[Publisher] [nvarchar](255) NULL,
	[StudioID] [nvarchar](255) NOT NULL,
	[VTSID] [nvarchar](255) NOT NULL
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[GamesTitles_RelDate]    Script Date: 1/27/2021 4:41:04 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[GamesTitles_RelDate](
	[UID] [int] IDENTITY(1,1) NOT NULL,
	[ID] [nvarchar](255) NOT NULL,
	[TitleID] [nvarchar](255) NULL,
	[TitleName] [nvarchar](255) NULL,
	[ReleasePlatform] [nvarchar](255) NULL,
	[ReleaseDate] [datetime] NULL,
	[IGDB_Website] [nvarchar](255) NULL,
	[NewZoo_Website] [nvarchar](255) NULL
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[GamesTitles_SupDev]    Script Date: 1/27/2021 4:41:04 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[GamesTitles_SupDev](
	[UID] [int] IDENTITY(1,1) NOT NULL,
	[ID] [nvarchar](255) NOT NULL,
	[TitleID] [nvarchar](255) NULL,
	[TitleName] [nvarchar](255) NULL,
	[IGDB_Website] [nvarchar](255) NULL,
	[NewZoo_Website] [nvarchar](255) NULL,
	[SupportDev] [nvarchar](255) NULL,
	[StudioID] [nvarchar](255) NOT NULL,
	[VTSID] [nvarchar](255) NOT NULL
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[Marketing_Audit]    Script Date: 1/27/2021 4:41:04 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[Marketing_Audit](
	[ID] [int] NULL,
	[Column_Name] [nvarchar](150) NULL,
	[Old_Value] [nvarchar](max) NULL,
	[New_Value] [nvarchar](max) NULL,
	[Activity] [nvarchar](20) NULL,
	[SourceTbl] [nvarchar](50) NULL,
	[Time_Stamp] [datetime] NOT NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

/****** Object:  Table [dbo].[Marketing_ETL]    Script Date: 1/27/2021 4:41:04 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[Marketing_ETL](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[VTSID] [nvarchar](255) NULL,
	[LinkedInID] [nvarchar](255) NULL,
	[CompanyName] [nvarchar](255) NULL,
	[Developer] [nvarchar](255) NULL,
	[SupportingDeveloper] [nvarchar](255) NULL,
	[PortingDeveloper] [nvarchar](255) NULL,
	[CompanyWebsite] [nvarchar](255) NULL,
	[EmployeeRange] [nvarchar](255) NULL,
	[UltimateParent] [nvarchar](255) NULL,
	[Parent] [nvarchar](255) NULL,
	[Subsidiaries] [nvarchar](max) NULL,
	[City] [nvarchar](255) NULL,
	[RegionStateProvince] [nvarchar](255) NULL,
	[Country] [nvarchar](255) NULL,
	[BusinessClassification] [nvarchar](255) NULL,
	[BusinessSubclassification] [nvarchar](255) NULL,
	[Active] [nvarchar](255) NULL,
	[Source] [nvarchar](255) NULL,
	[LinkedInURL] [nvarchar](255) NULL,
	[Description] [nvarchar](max) NULL,
	[Type] [nvarchar](255) NULL,
	[CompanyAddress] [nvarchar](255) NULL,
	[Phone] [nvarchar](255) NULL,
	[EmployeesonLinkedIn] [nvarchar](255) NULL,
	[Founded] [nvarchar](255) NULL,
	[Growth6mth] [float] NULL,
	[Growth1yr] [float] NULL,
	[Growth2yr] [float] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO


USE [Marketing]
GO

/****** Object:  Trigger [dbo].[Marketing_trigger_Update]    Script Date: 1/27/2021 4:41:18 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE trigger [dbo].[Marketing_trigger_Update]
on [dbo].[Marketing_ETL]
after UPDATE, INSERT, DELETE
as
DECLARE @ID int, @activity varchar(20), @newValue nvarchar(max),
        @oldValue nvarchar(max), @Column_Name nvarchar(100), @Source nvarchar(50);

SET @Source = 'Marketing_ETL'

BEGIN
	iF UPDATE(EmployeeRange)
	Begin
		SET @Column_Name = 'EmployeeRange'

		if exists(SELECT * from inserted) and exists (SELECT * from deleted)
		begin
			SET @activity = 'UPDATE';
			-- SET @user = SYSTEM_USER;
			-- SELECT @VTSID = VTSID from inserted i 
			SELECT @ID = inserted.ID, @newValue=inserted.EmployeeRange, @oldValue=deleted.EmployeeRange FROM inserted
			JOIN deleted ON inserted.ID = deleted.ID;
			if @newValue <> @oldvalue
			begin
				 INSERT into Marketing_Audit(ID, Column_Name, New_Value, Old_Value, Activity, SourceTbl, Time_Stamp) values (@ID, @Column_Name, @newValue, @oldValue, @activity, @Source, GETDATE());
			end
		end
	end

	iF UPDATE(UltimateParent)
	Begin
		SET @Column_Name = 'UltimateParent'

		if exists(SELECT * from inserted) and exists (SELECT * from deleted)
		begin
			SET @activity = 'UPDATE';
			-- SET @user = SYSTEM_USER;
			-- SELECT @VTSID = VTSID from inserted i 
			SELECT @ID = inserted.ID, @newValue=inserted.UltimateParent, @oldValue=deleted.UltimateParent FROM inserted
			JOIN deleted ON inserted.ID = deleted.ID;
			if @newValue <> @oldvalue
			begin
				 INSERT into Marketing_Audit(ID, Column_Name, New_Value, Old_Value, Activity, SourceTbl, Time_Stamp) values (@ID, @Column_Name, @newValue, @oldValue, @activity, @Source, GETDATE());
			end
		end
	end

	iF UPDATE(Parent)
	Begin
		SET @Column_Name = 'Parent'

		if exists(SELECT * from inserted) and exists (SELECT * from deleted)
		begin
			SET @activity = 'UPDATE';
			-- SET @user = SYSTEM_USER;
			-- SELECT @VTSID = VTSID from inserted i 
			SELECT @ID = inserted.ID, @newValue=inserted.Parent, @oldValue=deleted.Parent FROM inserted
			JOIN deleted ON inserted.ID = deleted.ID;
			if @newValue <> @oldvalue
			begin
				 INSERT into Marketing_Audit(ID, Column_Name, New_Value, Old_Value, Activity, SourceTbl, Time_Stamp) values (@ID, @Column_Name, @newValue, @oldValue, @activity, @Source, GETDATE());
			end
		end
	end

	iF UPDATE(Subsidiaries)
	Begin
		SET @Column_Name = 'Subsidiaries'

		if exists(SELECT * from inserted) and exists (SELECT * from deleted)
		begin
			SET @activity = 'UPDATE';
			-- SET @user = SYSTEM_USER;
			-- SELECT @VTSID = VTSID from inserted i 
			SELECT @ID = inserted.ID, @newValue=inserted.Subsidiaries, @oldValue=deleted.Subsidiaries FROM inserted
			JOIN deleted ON inserted.ID = deleted.ID;
			if @newValue <> @oldvalue
			begin
				 INSERT into Marketing_Audit(ID, Column_Name, New_Value, Old_Value, Activity, SourceTbl, Time_Stamp) values (@ID, @Column_Name, @newValue, @oldValue, @activity, @Source, GETDATE());
			end
		end
	end

	iF UPDATE(City)
	Begin
		SET @Column_Name = 'City'

		if exists(SELECT * from inserted) and exists (SELECT * from deleted)
		begin
			SET @activity = 'UPDATE';
			-- SET @user = SYSTEM_USER;
			-- SELECT @VTSID = VTSID from inserted i 
			SELECT @ID = inserted.ID, @newValue=inserted.City, @oldValue=deleted.City FROM inserted
			JOIN deleted ON inserted.ID = deleted.ID;
			if @newValue <> @oldvalue
			begin
				 INSERT into Marketing_Audit(ID, Column_Name, New_Value, Old_Value, Activity, SourceTbl, Time_Stamp) values (@ID, @Column_Name, @newValue, @oldValue, @activity, @Source, GETDATE());
			end
		end
	end

	iF UPDATE(RegionStateProvince)
	Begin
		SET @Column_Name = 'RegionStateProvince'

		if exists(SELECT * from inserted) and exists (SELECT * from deleted)
		begin
			SET @activity = 'UPDATE';
			-- SET @user = SYSTEM_USER;
			-- SELECT @VTSID = VTSID from inserted i 
			SELECT @ID = inserted.ID, @newValue=inserted.RegionStateProvince, @oldValue=deleted.RegionStateProvince FROM inserted
			JOIN deleted ON inserted.ID = deleted.ID;
			if @newValue <> @oldvalue
			begin
				 INSERT into Marketing_Audit(ID, Column_Name, New_Value, Old_Value, Activity, SourceTbl, Time_Stamp) values (@ID, @Column_Name, @newValue, @oldValue, @activity, @Source, GETDATE());
			end
		end
	end

	iF UPDATE(Country)
	Begin
		SET @Column_Name = 'Country'

		if exists(SELECT * from inserted) and exists (SELECT * from deleted)
		begin
			SET @activity = 'UPDATE';
			-- SET @user = SYSTEM_USER;
			-- SELECT @VTSID = VTSID from inserted i 
			SELECT @ID = inserted.ID, @newValue=inserted.Country, @oldValue=deleted.Country FROM inserted
			JOIN deleted ON inserted.ID = deleted.ID;
			if @newValue <> @oldvalue
			begin
				 INSERT into Marketing_Audit(ID, Column_Name, New_Value, Old_Value, Activity, SourceTbl, Time_Stamp) values (@ID, @Column_Name, @newValue, @oldValue, @activity, @Source, GETDATE());
			end
		end
	end

	iF UPDATE(BusinessClassification)
	Begin
		SET @Column_Name = 'BusinessClassification'

		if exists(SELECT * from inserted) and exists (SELECT * from deleted)
		begin
			SET @activity = 'UPDATE';
			-- SET @user = SYSTEM_USER;
			-- SELECT @VTSID = VTSID from inserted i 
			SELECT @ID = inserted.ID, @newValue=inserted.BusinessClassification, @oldValue=deleted.BusinessClassification FROM inserted
			JOIN deleted ON inserted.ID = deleted.ID;
			if @newValue <> @oldvalue
			begin
				 INSERT into Marketing_Audit(ID, Column_Name, New_Value, Old_Value, Activity, SourceTbl, Time_Stamp) values (@ID, @Column_Name, @newValue, @oldValue, @activity, @Source, GETDATE());
			end
		end
	end

	iF UPDATE(BusinessSubclassification)
	Begin
		SET @Column_Name = 'BusinessSubclassification'

		if exists(SELECT * from inserted) and exists (SELECT * from deleted)
		begin
			SET @activity = 'UPDATE';
			-- SET @user = SYSTEM_USER;
			-- SELECT @VTSID = VTSID from inserted i 
			SELECT @ID = inserted.ID, @newValue=inserted.BusinessSubclassification, @oldValue=deleted.BusinessSubclassification FROM inserted
			JOIN deleted ON inserted.ID = deleted.ID;
			if @newValue <> @oldvalue
			begin
				 INSERT into Marketing_Audit(ID, Column_Name, New_Value, Old_Value, Activity, SourceTbl, Time_Stamp) values (@ID, @Column_Name, @newValue, @oldValue, @activity, @Source, GETDATE());
			end
		end
	end

	iF UPDATE(Active)
	Begin
		SET @Column_Name = 'Active'

		if exists(SELECT * from inserted) and exists (SELECT * from deleted)
		begin
			SET @activity = 'UPDATE';
			-- SET @user = SYSTEM_USER;
			-- SELECT @VTSID = VTSID from inserted i 
			SELECT @ID = inserted.ID, @newValue=inserted.Active, @oldValue=deleted.Active FROM inserted
			JOIN deleted ON inserted.ID = deleted.ID;
			if @newValue <> @oldvalue
			begin
				 INSERT into Marketing_Audit(ID, Column_Name, New_Value, Old_Value, Activity, SourceTbl, Time_Stamp) values (@ID, @Column_Name, @newValue, @oldValue, @activity, @Source, GETDATE());
			end
		end
	end

	iF UPDATE(Source)
	Begin
		SET @Column_Name = 'Source'

		if exists(SELECT * from inserted) and exists (SELECT * from deleted)
		begin
			SET @activity = 'UPDATE';
			-- SET @user = SYSTEM_USER;
			-- SELECT @VTSID = VTSID from inserted i 
			SELECT @ID = inserted.ID, @newValue=inserted.Source, @oldValue=deleted.Source FROM inserted
			JOIN deleted ON inserted.ID = deleted.ID;
			if @newValue <> @oldvalue
			begin
				 INSERT into Marketing_Audit(ID, Column_Name, New_Value, Old_Value, Activity, SourceTbl, Time_Stamp) values (@ID, @Column_Name, @newValue, @oldValue, @activity, @Source, GETDATE());
			end
		end
	end

	iF UPDATE(LinkedInURL)
	Begin
		SET @Column_Name = 'LinkedInURL'

		if exists(SELECT * from inserted) and exists (SELECT * from deleted)
		begin
			SET @activity = 'UPDATE';
			-- SET @user = SYSTEM_USER;
			-- SELECT @VTSID = VTSID from inserted i 
			SELECT @ID = inserted.ID, @newValue=inserted.Source, @oldValue=deleted.Source FROM inserted
			JOIN deleted ON inserted.ID = deleted.ID;
			if @newValue <> @oldvalue
			begin
				 INSERT into Marketing_Audit(ID, Column_Name, New_Value, Old_Value, Activity, SourceTbl, Time_Stamp) values (@ID, @Column_Name, @newValue, @oldValue, @activity, @Source, GETDATE());
			end
		end
	end


	iF UPDATE(EmployeesonLinkedIn)
	Begin
		SET @Column_Name = 'EmployeesonLinkedIn'

		if exists(SELECT * from inserted) and exists (SELECT * from deleted)
		begin
			SET @activity = 'UPDATE';
			-- SET @user = SYSTEM_USER;
			-- SELECT @VTSID = VTSID from inserted i 
			SELECT @ID = inserted.ID, @newValue=inserted.EmployeesonLinkedIn, @oldValue=deleted.EmployeesonLinkedIn FROM inserted
			JOIN deleted ON inserted.ID = deleted.ID;
			if @newValue <> @oldvalue
			begin
				 INSERT into Marketing_Audit(ID, Column_Name, New_Value, Old_Value, Activity, SourceTbl, Time_Stamp) values (@ID, @Column_Name, @newValue, @oldValue, @activity, @Source, GETDATE());
			end
		end
	end

	iF UPDATE(Growth6mth)
	Begin
		SET @Column_Name = 'Growth6mth'

		if exists(SELECT * from inserted) and exists (SELECT * from deleted)
		begin
			SET @activity = 'UPDATE';
			-- SET @user = SYSTEM_USER;
			-- SELECT @VTSID = VTSID from inserted i 
			SELECT @ID = inserted.ID, @newValue=inserted.Growth6mth, @oldValue=deleted.Growth6mth FROM inserted
			JOIN deleted ON inserted.ID = deleted.ID;
			if @newValue <> @oldvalue
			begin
				 INSERT into Marketing_Audit(ID, Column_Name, New_Value, Old_Value, Activity, SourceTbl, Time_Stamp) values (@ID, @Column_Name, @newValue, @oldValue, @activity, @Source, GETDATE());
			end
		end
	end

	iF UPDATE(Growth1yr)
	Begin
		SET @Column_Name = 'Growth1yr'

		if exists(SELECT * from inserted) and exists (SELECT * from deleted)
		begin
			SET @activity = 'UPDATE';
			-- SET @user = SYSTEM_USER;
			-- SELECT @VTSID = VTSID from inserted i 
			SELECT @ID = inserted.ID, @newValue=inserted.Growth1yr, @oldValue=deleted.Growth1yr FROM inserted
			JOIN deleted ON inserted.ID = deleted.ID;
			if @newValue <> @oldvalue
			begin
				 INSERT into Marketing_Audit(ID, Column_Name, New_Value, Old_Value, Activity, SourceTbl, Time_Stamp) values (@ID, @Column_Name, @newValue, @oldValue, @activity, @Source, GETDATE());
			end
		end
	end

	iF UPDATE(Growth2yr)
	Begin
		SET @Column_Name = 'Growth2yr'

		if exists(SELECT * from inserted) and exists (SELECT * from deleted)
		begin
			SET @activity = 'UPDATE';
			-- SET @user = SYSTEM_USER;
			-- SELECT @VTSID = VTSID from inserted i 
			SELECT @ID = inserted.ID, @newValue=inserted.Growth2yr, @oldValue=deleted.Growth2yr FROM inserted
			JOIN deleted ON inserted.ID = deleted.ID;
			if @newValue <> @oldvalue
			begin
				 INSERT into Marketing_Audit(ID, Column_Name, New_Value, Old_Value, Activity, SourceTbl, Time_Stamp) values (@ID, @Column_Name, @newValue, @oldValue, @activity, @Source, GETDATE());
			end
		end
	end

END
GO

ALTER TABLE [dbo].[Marketing_ETL] ENABLE TRIGGER [Marketing_trigger_Update]
GO


USE [Marketing]
GO

/****** Object:  Trigger [dbo].[GameTitles_trigger_Update]    Script Date: 1/27/2021 4:41:43 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE trigger [dbo].[GameTitles_trigger_Update]
on [dbo].[GamesTitles]
after UPDATE, INSERT, DELETE
as
DECLARE @ID nvarchar(255), @activity varchar(20), @newValue nvarchar(max),
        @oldValue nvarchar(max), @Column_Name nvarchar(100), @Source nvarchar(50);

SET @Source = 'GamesTitle'

BEGIN
	iF UPDATE(TitleName)
	Begin
		SET @Column_Name = 'TitleName'

		if exists(SELECT * from inserted) and exists (SELECT * from deleted)
		begin
			SET @activity = 'UPDATE';
			-- SET @user = SYSTEM_USER;
			-- SELECT @VTSID = VTSID from inserted i 
			SELECT @ID = inserted.ID, @newValue=inserted.TitleName, @oldValue=deleted.TitleName FROM inserted
			JOIN deleted ON inserted.ID = deleted.ID;
			if @newValue <> @oldvalue
			begin
				 INSERT into Marketing_Audit(ID, Column_Name, New_Value, Old_Value, Activity, SourceTbl, Time_Stamp) values (@ID, @Column_Name, @newValue, @oldValue, @activity, @Source, GETDATE());
			end
		end
	end

	iF UPDATE(Metascore)
	Begin
		SET @Column_Name = 'Metascore'

		if exists(SELECT * from inserted) and exists (SELECT * from deleted)
		begin
			SET @activity = 'UPDATE';
			-- SET @user = SYSTEM_USER;
			-- SELECT @VTSID = VTSID from inserted i 
			SELECT @ID = inserted.ID, @newValue=inserted.Metascore, @oldValue=deleted.Metascore FROM inserted
			JOIN deleted ON inserted.ID = deleted.ID;
			if @newValue <> @oldvalue
			begin
				 INSERT into Marketing_Audit(ID, Column_Name, New_Value, Old_Value, Activity, SourceTbl, Time_Stamp) values (@ID, @Column_Name, @newValue, @oldValue, @activity, @Source, GETDATE());
			end
		end
	end

	iF UPDATE(GameModes)
	Begin
		SET @Column_Name = 'GameModes'

		if exists(SELECT * from inserted) and exists (SELECT * from deleted)
		begin
			SET @activity = 'UPDATE';
			-- SET @user = SYSTEM_USER;
			-- SELECT @VTSID = VTSID from inserted i 
			SELECT @ID = inserted.ID, @newValue=inserted.GameModes, @oldValue=deleted.GameModes FROM inserted
			JOIN deleted ON inserted.ID = deleted.ID;
			if @newValue <> @oldvalue
			begin
				 INSERT into Marketing_Audit(ID, Column_Name, New_Value, Old_Value, Activity, SourceTbl, Time_Stamp) values (@ID, @Column_Name, @newValue, @oldValue, @activity, @Source, GETDATE());
			end
		end
	end

	iF UPDATE(Genre)
	Begin
		SET @Column_Name = 'Genre'

		if exists(SELECT * from inserted) and exists (SELECT * from deleted)
		begin
			SET @activity = 'UPDATE';
			-- SET @user = SYSTEM_USER;
			-- SELECT @VTSID = VTSID from inserted i 
			SELECT @ID = inserted.ID, @newValue=inserted.Genre, @oldValue=deleted.Genre FROM inserted
			JOIN deleted ON inserted.ID = deleted.ID;
			if @newValue <> @oldvalue
			begin
				 INSERT into Marketing_Audit(ID, Column_Name, New_Value, Old_Value, Activity, SourceTbl, Time_Stamp) values (@ID, @Column_Name, @newValue, @oldValue, @activity, @Source, GETDATE());
			end
		end
	end

	iF UPDATE(Themes)
	Begin
		SET @Column_Name = 'Themes'

		if exists(SELECT * from inserted) and exists (SELECT * from deleted)
		begin
			SET @activity = 'UPDATE';
			-- SET @user = SYSTEM_USER;
			-- SELECT @VTSID = VTSID from inserted i 
			SELECT @ID = inserted.ID, @newValue=inserted.Themes, @oldValue=deleted.Themes FROM inserted
			JOIN deleted ON inserted.ID = deleted.ID;
			if @newValue <> @oldvalue
			begin
				 INSERT into Marketing_Audit(ID, Column_Name, New_Value, Old_Value, Activity, SourceTbl, Time_Stamp) values (@ID, @Column_Name, @newValue, @oldValue, @activity, @Source, GETDATE());
			end
		end
	end

	iF UPDATE(Series)
	Begin
		SET @Column_Name = 'Series'

		if exists(SELECT * from inserted) and exists (SELECT * from deleted)
		begin
			SET @activity = 'UPDATE';
			-- SET @user = SYSTEM_USER;
			-- SELECT @VTSID = VTSID from inserted i 
			SELECT @ID = inserted.ID, @newValue=inserted.Series, @oldValue=deleted.Series FROM inserted
			JOIN deleted ON inserted.ID = deleted.ID;
			if @newValue <> @oldvalue
			begin
				 INSERT into Marketing_Audit(ID, Column_Name, New_Value, Old_Value, Activity, SourceTbl, Time_Stamp) values (@ID, @Column_Name, @newValue, @oldValue, @activity, @Source, GETDATE());
			end
		end
	end

	iF UPDATE(PlayerPerspectives)
	Begin
		SET @Column_Name = 'PlayerPerspectives'

		if exists(SELECT * from inserted) and exists (SELECT * from deleted)
		begin
			SET @activity = 'UPDATE';
			-- SET @user = SYSTEM_USER;
			-- SELECT @VTSID = VTSID from inserted i 
			SELECT @ID = inserted.ID, @newValue=inserted.PlayerPerspectives, @oldValue=deleted.PlayerPerspectives FROM inserted
			JOIN deleted ON inserted.ID = deleted.ID;
			if @newValue <> @oldvalue
			begin
				 INSERT into Marketing_Audit(ID, Column_Name, New_Value, Old_Value, Activity, SourceTbl, Time_Stamp) values (@ID, @Column_Name, @newValue, @oldValue, @activity, @Source, GETDATE());
			end
		end
	end

	iF UPDATE(Franchises)
	Begin
		SET @Column_Name = 'Franchises'

		if exists(SELECT * from inserted) and exists (SELECT * from deleted)
		begin
			SET @activity = 'UPDATE';
			-- SET @user = SYSTEM_USER;
			-- SELECT @VTSID = VTSID from inserted i 
			SELECT @ID = inserted.ID, @newValue=inserted.Franchises, @oldValue=deleted.Franchises FROM inserted
			JOIN deleted ON inserted.ID = deleted.ID;
			if @newValue <> @oldvalue
			begin
				 INSERT into Marketing_Audit(ID, Column_Name, New_Value, Old_Value, Activity, SourceTbl, Time_Stamp) values (@ID, @Column_Name, @newValue, @oldValue, @activity, @Source, GETDATE());
			end
		end
	end

	iF UPDATE(GameEngine)
	Begin
		SET @Column_Name = 'GameEngine'

		if exists(SELECT * from inserted) and exists (SELECT * from deleted)
		begin
			SET @activity = 'UPDATE';
			-- SET @user = SYSTEM_USER;
			-- SELECT @VTSID = VTSID from inserted i 
			SELECT @ID = inserted.ID, @newValue=inserted.GameEngine, @oldValue=deleted.GameEngine FROM inserted
			JOIN deleted ON inserted.ID = deleted.ID;
			if @newValue <> @oldvalue
			begin
				 INSERT into Marketing_Audit(ID, Column_Name, New_Value, Old_Value, Activity, SourceTbl, Time_Stamp) values (@ID, @Column_Name, @newValue, @oldValue, @activity, @Source, GETDATE());
			end
		end
	end

	iF UPDATE(AlternativeNames)
	Begin
		SET @Column_Name = 'AlternativeNames'

		if exists(SELECT * from inserted) and exists (SELECT * from deleted)
		begin
			SET @activity = 'UPDATE';
			-- SET @user = SYSTEM_USER;
			-- SELECT @VTSID = VTSID from inserted i 
			SELECT @ID = inserted.ID, @newValue=inserted.AlternativeNames, @oldValue=deleted.AlternativeNames FROM inserted
			JOIN deleted ON inserted.ID = deleted.ID;
			if @newValue <> @oldvalue
			begin
				 INSERT into Marketing_Audit(ID, Column_Name, New_Value, Old_Value, Activity, SourceTbl, Time_Stamp) values (@ID, @Column_Name, @newValue, @oldValue, @activity, @Source, GETDATE());
			end
		end
	end

	iF UPDATE(IGDB_Website)
	Begin
		SET @Column_Name = 'IGDB_Website'

		if exists(SELECT * from inserted) and exists (SELECT * from deleted)
		begin
			SET @activity = 'UPDATE';
			-- SET @user = SYSTEM_USER;
			-- SELECT @VTSID = VTSID from inserted i 
			SELECT @ID = inserted.ID, @newValue=inserted.IGDB_Website, @oldValue=deleted.IGDB_Website FROM inserted
			JOIN deleted ON inserted.ID = deleted.ID;
			if @newValue <> @oldvalue
			begin
				 INSERT into Marketing_Audit(ID, Column_Name, New_Value, Old_Value, Activity, SourceTbl, Time_Stamp) values (@ID, @Column_Name, @newValue, @oldValue, @activity, @Source, GETDATE());
			end
		end
	end

	iF UPDATE(NewZoo_Website)
	Begin
		SET @Column_Name = 'NewZoo_Website'

		if exists(SELECT * from inserted) and exists (SELECT * from deleted)
		begin
			SET @activity = 'UPDATE';
			-- SET @user = SYSTEM_USER;
			-- SELECT @VTSID = VTSID from inserted i 
			SELECT @ID = inserted.ID, @newValue=inserted.NewZoo_Website, @oldValue=deleted.NewZoo_Website FROM inserted
			JOIN deleted ON inserted.ID = deleted.ID;
			if @newValue <> @oldvalue
			begin
				 INSERT into Marketing_Audit(ID, Column_Name, New_Value, Old_Value, Activity, SourceTbl, Time_Stamp) values (@ID, @Column_Name, @newValue, @oldValue, @activity, @Source, GETDATE());
			end
		end
	end

END
GO

ALTER TABLE [dbo].[GamesTitles] ENABLE TRIGGER [GameTitles_trigger_Update]
GO


USE [Marketing]
GO

/****** Object:  Trigger [dbo].[GameTitles_RelDate_trigger_Update]    Script Date: 1/27/2021 4:42:09 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE trigger [dbo].[GameTitles_RelDate_trigger_Update]
on [dbo].[GamesTitles_RelDate]
after UPDATE, INSERT, DELETE
as
DECLARE @UID nvarchar(255), @activity varchar(20), @newValue nvarchar(50),
        @oldValue nvarchar(50), @Column_Name nvarchar(100), @Source nvarchar(50);

SET @Source = 'GamesTitle_RelDate'

BEGIN
	iF UPDATE(ReleaseDate)
	Begin
		SET @Column_Name = 'ReleaseDate'

		if exists(SELECT * from inserted) and exists (SELECT * from deleted)
		begin
			SET @activity = 'UPDATE';
			-- SET @user = SYSTEM_USER;
			-- SELECT @VTSID = VTSID from inserted i 
			SELECT @UID = inserted.UID, @newValue=convert(varchar, CONVERT(datetime, inserted.ReleaseDate), 21), @oldValue=convert(varchar, CONVERT(datetime, deleted.ReleaseDate), 21) FROM inserted
			JOIN deleted ON inserted.UID = deleted.UID;
			if @newValue <> @oldvalue
			begin
				 INSERT into Marketing_Audit(ID, Column_Name, New_Value, Old_Value, Activity, SourceTbl, Time_Stamp) values (@UID, @Column_Name, @newValue, @oldValue, @activity, @Source, GETDATE());
			end
		end
	end

END
GO

ALTER TABLE [dbo].[GamesTitles_RelDate] ENABLE TRIGGER [GameTitles_RelDate_trigger_Update]
GO


USE [Marketing]
GO

/****** Object:  StoredProcedure [dbo].[UPDATE_Title_StudioID]    Script Date: 1/27/2021 4:42:44 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[UPDATE_Title_StudioID] 
AS

UPDATE [dbo].[GamesTitles_Dev] SET StudioID=C.ID
FROM [dbo].[GamesTitles_Dev] A INNER JOIN
(
	SELECT A.UID, A.VTSID, B.ID FROM 
	(SELECT UID, VTSID FROM [dbo].[GamesTitles_Dev] WHERE VTSID<>'') A
	INNER JOIN
	(SELECT ID, VTSID FROM [dbo].[Marketing_ETL]) B
	ON A.VTSID=B.VTSID
) C ON A.UID=C.UID WHERE A.VTSID<>''


UPDATE [dbo].[GamesTitles_PortDev] SET StudioID=C.ID
FROM [dbo].[GamesTitles_PortDev] A INNER JOIN
(
	SELECT A.UID, A.VTSID, B.ID FROM 
	(SELECT UID, VTSID FROM [dbo].[GamesTitles_PortDev] WHERE VTSID<>'') A
	INNER JOIN
	(SELECT ID, VTSID FROM [dbo].[Marketing_ETL]) B
	ON A.VTSID=B.VTSID
) C ON A.UID=C.UID WHERE A.VTSID<>''


UPDATE [dbo].[GamesTitles_Pub] SET StudioID=C.ID
FROM [dbo].[GamesTitles_Pub] A INNER JOIN
(
	SELECT A.UID, A.VTSID, B.ID FROM 
	(SELECT UID, VTSID FROM [dbo].[GamesTitles_Pub] WHERE VTSID<>'') A
	INNER JOIN
	(SELECT ID, VTSID FROM [dbo].[Marketing_ETL]) B
	ON A.VTSID=B.VTSID
) C ON A.UID=C.UID WHERE A.VTSID<>''


UPDATE [dbo].[GamesTitles_SupDev] SET StudioID=C.ID
FROM [dbo].[GamesTitles_SupDev] A INNER JOIN
(
	SELECT A.UID, A.VTSID, B.ID FROM 
	(SELECT UID, VTSID FROM [dbo].[GamesTitles_SupDev] WHERE VTSID<>'') A
	INNER JOIN
	(SELECT ID, VTSID FROM [dbo].[Marketing_ETL]) B
	ON A.VTSID=B.VTSID
) C ON A.UID=C.UID WHERE A.VTSID<>''

GO


