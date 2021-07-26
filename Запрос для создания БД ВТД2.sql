DROP TABLE [VTD].[pipelineInfo] 
DROP TABLE [VTD].[furnishingsLog]
DROP TABLE [VTD].[anomalyLogLine]
DROP TABLE [VTD].[MGPipe]  
GO
CREATE SCHEMA [VTD] AUTHORIZATION [dbo]; 
GO  
PRINT 'SCHEMA "VTD" is created';  
GO 


CREATE TABLE [VTD].[pipelineInfo] (  
    [itemNumber]   INT           IDENTITY (1, 1) NOT NULL,  
    [pipelineName]   NVARCHAR (100)              NOT NULL, 
    [pipelineSection]   NVARCHAR (100)           NOT NULL, 
    [pipeDiameter]  DECIMAL                      NOT NULL,
    [principal]   NVARCHAR (100)                 NOT NULL,
    [examinationDate]   NVARCHAR (100)           NOT NULL,
    [designPressure]   DECIMAL                   NOT NULL,
    [operatingPressure]   DECIMAL                NOT NULL,
    [comissioningYear]   NVARCHAR (100)          NOT NULL,
    );  
    GO  
    PRINT 'Table "pipelineInfo" is created';  
    GO 
    ALTER TABLE [VTD].[pipelineInfo]  
    ADD CONSTRAINT [Defolt_zero_for_all_pipelineInfo] DEFAULT 0 FOR [pipelineName];
    ALTER TABLE [VTD].[pipelineInfo]  
    ADD CONSTRAINT [Defolt_zero_for_all_pipelineInfo] DEFAULT 0 FOR [pipelineSection];
    ALTER TABLE [VTD].[pipelineInfo]  
    ADD CONSTRAINT [Defolt_zero_for_all_pipelineInfo] DEFAULT 0 FOR [pipeDiameter];
    ALTER TABLE [VTD].[pipelineInfo]  
    ADD CONSTRAINT [Defolt_zero_for_all_pipelineInfo] DEFAULT 0 FOR [principal];
    ALTER TABLE [VTD].[pipelineInfo]  
    ADD CONSTRAINT [Defolt_zero_for_all_pipelineInfo] DEFAULT 0 FOR [examinationDate];
    ALTER TABLE [VTD].[pipelineInfo]  
    ADD CONSTRAINT [Defolt_zero_for_all_pipelineInfo] DEFAULT 0 FOR [designPressure];
    ALTER TABLE [VTD].[pipelineInfo]  
    ADD CONSTRAINT [Defolt_zero_for_all_pipelineInfo] DEFAULT 0 FOR [ operatingPressure];
    ALTER TABLE [VTD].[pipelineInfo]  
    ADD CONSTRAINT [Defolt_zero_for_all_pipelineInfo] DEFAULT 0 FOR [ comissioningYear];
    GO  
    PRINT 'CONSTRAINT "Defolt_zero_for_all_pipelineInfo" for table [VTD].[pipelineInfo] is created';  
    GO 
    CREATE TABLE [VTD].[MGPipe] (  
    [itemNumber]   INT           IDENTITY (1, 1) NOT NULL,  
    [pipeNumber]   NVARCHAR (100)                NOT NULL, 
    [odometrDist]  DECIMAL                       NOT NULL,
    [pipeLength]  DECIMAL                        NOT NULL,
    [distanceFromReferencePoints]  DECIMAL       NOT NULL,
    [characterFeatures]   NVARCHAR (100)         NOT NULL, 
    [clockOrientation]   NVARCHAR (100)          NOT NULL, 
    [bendOfPipe]  DECIMAL                        NOT NULL,
    [jointAngle]  DECIMAL                        NOT NULL,
    [Latitude]   VARCHAR (20)                    NOT NULL,
    [Longitude]   VARCHAR (20)                   NOT NULL,
    [heightAboveSeaLevel]  DECIMAL               NOT NULL,
    [note]   NVARCHAR (100)                      NOT NULL,   
    );  
    GO  
    PRINT 'Table "MGPipe" is created';  
    GO 
    ALTER TABLE [VTD].[MGPipe]  
    ADD CONSTRAINT [Defolt_zero_for_all_pipelineInfo] DEFAULT 0 FOR [pipeNumber];
    ALTER TABLE [VTD].[MGPipe]  
    ADD CONSTRAINT [Defolt_zero_for_all_pipelineInfo] DEFAULT 0 FOR [odometrDist];
    ALTER TABLE [VTD].[MGPipe]  
    ADD CONSTRAINT [Defolt_zero_for_all_pipelineInfo] DEFAULT 0 FOR [ pipeLength];
    ALTER TABLE [VTD].[MGPipe]  
    ADD CONSTRAINT [Defolt_zero_for_all_pipelineInfo] DEFAULT 0 FOR [distanceFromReferencePoints];
    ALTER TABLE [VTD].[MGPipe]  
    ADD CONSTRAINT [Defolt_zero_for_all_pipelineInfo] DEFAULT 0 FOR [characterFeatures];
    ALTER TABLE [VTD].[MGPipe]  
    ADD CONSTRAINT [Defolt_zero_for_all_pipelineInfo] DEFAULT 0 FOR [clockOrientation];
    ALTER TABLE [VTD].[MGPipe]  
    ADD CONSTRAINT [Defolt_zero_for_all_pipelineInfo] DEFAULT 0 FOR [bendOfPipe];
    ALTER TABLE [VTD].[MGPipe]  
    ADD CONSTRAINT [Defolt_zero_for_all_pipelineInfo] DEFAULT 0 FOR [jointAngle];
    ALTER TABLE [VTD].[MGPipe]  
    ADD CONSTRAINT [Defolt_zero_for_all_pipelineInfo] DEFAULT 0 FOR [Latitude];
    ALTER TABLE [VTD].[MGPipe]  
    ADD CONSTRAINT [Defolt_zero_for_all_pipelineInfo] DEFAULT 0 FOR [Longitude];
    ALTER TABLE [VTD].[MGPipe]  
    ADD CONSTRAINT [Defolt_zero_for_all_pipelineInfo] DEFAULT 0 FOR [heightAboveSeaLevel];
    ALTER TABLE [VTD].[MGPipe]  
    ADD CONSTRAINT [Defolt_zero_for_all_pipelineInfo] DEFAULT 0 FOR [note];

    GO  
    PRINT 'CONSTRAINT "Defolt_zero_for_all_MGPipe" for table [VTD].[MGPipe] is created';  
    GO 

     CREATE TABLE [VTD].[anomalyLogLine] (  
    [itemNumber]   INT           IDENTITY (1, 1) NOT NULL,       
    [odometrDist]  DECIMAL                       NOT NULL,
    [distanceFromTransverseWeld]  DECIMAL        NOT NULL,
    [distanceFromReferencePoints]  DECIMAL       NOT NULL,
    [featuresCharacter]   NVARCHAR (100)         NOT NULL,
    [classOfSize]   NVARCHAR (100)               NOT NULL,
    [featuresOreientation]   NVARCHAR (100)      NOT NULL,
    [length]  DECIMAL                            NOT NULL,
    [widht]  DECIMAL                             NOT NULL,
    [depthInProcent]  DECIMAL                    NOT NULL,
    [depthInMm]  DECIMAL                         NOT NULL,
    [extOrInt]   NVARCHAR (100)                  NOT NULL,
    [KBD]  DECIMAL                               NOT NULL,
    [defectAssessment]   NVARCHAR (100)          NOT NULL,
    [Latitude]   VARCHAR (20)                    NOT NULL,
    [Longitude]   VARCHAR (20)                   NOT NULL,
    [heightAboveSeaLevel]  DECIMAL               NOT NULL,
    [note]   NVARCHAR (100)                      NOT NULL,   
    );  
    
    GO  
    PRINT 'Table "anomalyLogLine" is created';  
    GO 
    ALTER TABLE [VTD].[anomalyLogLine]  
    ADD CONSTRAINT [Defolt_zero_for_all_anomalyLogLine] DEFAULT 0 FOR [odometrDist];
    ALTER TABLE [VTD].[anomalyLogLine]  
    ADD CONSTRAINT [Defolt_zero_for_all_anomalyLogLine] DEFAULT 0 FOR [distanceFromTransverseWeld];
    ALTER TABLE [VTD].[anomalyLogLine]  
    ADD CONSTRAINT [Defolt_zero_for_all_anomalyLogLine] DEFAULT 0 FOR [distanceFromReferencePoints];
    ALTER TABLE [VTD].[anomalyLogLine]  
    ADD CONSTRAINT [Defolt_zero_for_all_anomalyLogLine] DEFAULT 0 FOR [featuresCharacter];
    ALTER TABLE [VTD].[anomalyLogLine]  
    ADD CONSTRAINT [Defolt_zero_for_all_anomalyLogLine] DEFAULT 0 FOR [classOfSize];
    ALTER TABLE [VTD].[anomalyLogLine]  
    ADD CONSTRAINT [Defolt_zero_for_all_anomalyLogLine] DEFAULT 0 FOR [featuresOreientation];
    ALTER TABLE [VTD].[anomalyLogLine]  
    ADD CONSTRAINT [Defolt_zero_for_all_anomalyLogLine] DEFAULT 0 FOR [length];
    ALTER TABLE [VTD].[anomalyLogLine]  
    ADD CONSTRAINT [Defolt_zero_for_all_anomalyLogLine] DEFAULT 0 FOR [widht];
    ALTER TABLE [VTD].[anomalyLogLine]  
    ADD CONSTRAINT [Defolt_zero_for_all_anomalyLogLine] DEFAULT 0 FOR [depthInProcent];
    ALTER TABLE [VTD].[anomalyLogLine]  
    ADD CONSTRAINT [Defolt_zero_for_all_anomalyLogLine] DEFAULT 0 FOR [depthInMm];
    ALTER TABLE [VTD].[anomalyLogLine]  
    ADD CONSTRAINT [Defolt_zero_for_all_anomalyLogLine] DEFAULT 0 FOR [extOrInt];
    ALTER TABLE [VTD].[anomalyLogLine]  
    ADD CONSTRAINT [Defolt_zero_for_all_anomalyLogLine] DEFAULT 0 FOR [KBD];
    ALTER TABLE [VTD].[anomalyLogLine]  
    ADD CONSTRAINT [Defolt_zero_for_all_anomalyLogLine] DEFAULT 0 FOR [defectAssessment];
    ALTER TABLE [VTD].[anomalyLogLine]  
    ADD CONSTRAINT [Defolt_zero_for_all_anomalyLogLine] DEFAULT 0 FOR [Latitude];
    ALTER TABLE [VTD].[anomalyLogLine]  
    ADD CONSTRAINT [Defolt_zero_for_all_anomalyLogLine] DEFAULT 0 FOR [Longitude];
    ALTER TABLE [VTD].[anomalyLogLine]  
    ADD CONSTRAINT [Defolt_zero_for_all_anomalyLogLine] DEFAULT 0 FOR [heightAboveSeaLevel];
    ALTER TABLE [VTD].[anomalyLogLine]  
    ADD CONSTRAINT [Defolt_zero_for_all_anomalyLogLine] DEFAULT 0 FOR [note];

    GO  
    PRINT 'CONSTRAINT "Defolt_zero_for_all_anomalyLogLine" for table [VTD].[anomalyLogLine] is created';  
    GO 

CREATE TABLE [VTD].[furnishingsLog] (  
    [itemNumber]   INT           IDENTITY (1, 1) NOT NULL,  
    [pipeNumber]   NVARCHAR (100)                NOT NULL, 
    [odometrDist]  DECIMAL                       NOT NULL,
    [pipeLength]  DECIMAL                        NOT NULL,
    [distanceFromTransverseWeld]  DECIMAL        NOT NULL,
    [characterFeatures]   NVARCHAR (100)         NOT NULL, 
    [designations]   NVARCHAR (100)              NOT NULL, 
    [marker]   NVARCHAR (100)                    NOT NULL, 
    [distanceToNextFeature]  DECIMAL             NOT NULL,
    [Latitude]   VARCHAR (20)                    NOT NULL,
    [Longitude]   VARCHAR (20)                   NOT NULL,
    [heightAboveSeaLevel]  DECIMAL               NOT NULL,
    [note]   NVARCHAR (100)                      NOT NULL,
    );  
    GO  
    PRINT 'Table "furnishingsLog" is created';  
    GO 
    ALTER TABLE [VTD].[furnishingsLog]  
    ADD CONSTRAINT [Defolt_zero_for_all_furnishingsLog] DEFAULT 0 FOR [pipeNumber];
        ALTER TABLE [VTD].[furnishingsLog]  
    ADD CONSTRAINT [Defolt_zero_for_all_furnishingsLog] DEFAULT 0 FOR [odometrDist];
        ALTER TABLE [VTD].[furnishingsLog]  
    ADD CONSTRAINT [Defolt_zero_for_all_furnishingsLog] DEFAULT 0 FOR [pipeLength];
        ALTER TABLE [VTD].[furnishingsLog]  
    ADD CONSTRAINT [Defolt_zero_for_all_furnishingsLog] DEFAULT 0 FOR [distanceFromTransverseWeld];
        ALTER TABLE [VTD].[furnishingsLog]  
    ADD CONSTRAINT [Defolt_zero_for_all_furnishingsLog] DEFAULT 0 FOR [characterFeatures];
        ALTER TABLE [VTD].[furnishingsLog]  
    ADD CONSTRAINT [Defolt_zero_for_all_furnishingsLog] DEFAULT 0 FOR [designations];
        ALTER TABLE [VTD].[furnishingsLog]  
    ADD CONSTRAINT [Defolt_zero_for_all_furnishingsLog] DEFAULT 0 FOR [marker];
        ALTER TABLE [VTD].[furnishingsLog]  
    ADD CONSTRAINT [Defolt_zero_for_all_furnishingsLog] DEFAULT 0 FOR [distanceToNextFeature];
        ALTER TABLE [VTD].[furnishingsLog]  
    ADD CONSTRAINT [Defolt_zero_for_all_furnishingsLog] DEFAULT 0 FOR [Latitude];
        ALTER TABLE [VTD].[furnishingsLog]  
    ADD CONSTRAINT [Defolt_zero_for_all_furnishingsLog] DEFAULT 0 FOR [Longitude];
        ALTER TABLE [VTD].[furnishingsLog]  
    ADD CONSTRAINT [Defolt_zero_for_all_furnishingsLog] DEFAULT 0 FOR [heightAboveSeaLevel];
        ALTER TABLE [VTD].[furnishingsLog]  
    ADD CONSTRAINT [Defolt_zero_for_all_furnishingsLog] DEFAULT 0 FOR [note];
    GO  
    PRINT 'CONSTRAINT "Defolt_zero_for_all_furnishingsLog" for table [VTD].[furnishingsLog] is created';  
    GO 
    --INSERT INTO [VTD].[furnishingsLog] (pipeNumber, odometrDist, pipeLength, distanceFromTransverseWeld, 
    --characterFeatures, designations, marker, distanceToNextFeature, Latitude, Longitude, heightAboveSeaLevel, note) values (555,555,555,555,555,555,555,555,555,555,555,555)