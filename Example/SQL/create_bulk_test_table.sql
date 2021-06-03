
CREATE TABLE dbo.bulk_test
(
    ID      int         NOT NULL,
    city    varchar(60) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
    country varchar(60) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
    upload_session_id  varchar(60) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
    create_app_user_id varchar(60) COLLATE SQL_Latin1_General_CP1_CI_AS NULL

)
go
IF OBJECT_ID(N'dbo.bulk_test') IS NOT NULL
    PRINT N'<<< CREATED TABLE dbo.bulk_test >>>'
ELSE
    PRINT N'<<< FAILED CREATING TABLE dbo.bulk_test >>>'
go
