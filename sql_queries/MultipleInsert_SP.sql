GO

CREATE PROCEDURE MultipleInsert( 
@FacultyAbbreviation nvarchar(max), 
@FacultyName nvarchar(max),
@FacultyWebsite nvarchar(max),
@FacultyEmail nvarchar(max),
@FacultyPhone nvarchar(max),
@UniversityId int,
@FacultyId int output,
--
@DepartmentAbbreviation nvarchar(max), 
@DepartmentName nvarchar(max),
@DepartmentWebsite nvarchar(max),
@DepartmentEmail nvarchar(max),
@DepartmentPhone nvarchar(max),
@DepartmentId int output
)
AS
BEGIN
	SET NOCOUNT ON;

	BEGIN TRANSACTION;
	SAVE TRANSACTION MySavePoint;

    BEGIN TRY

        INSERT INTO [dbo].[Faculties]
            (
            [Abbreviation]
            ,[Name]
            ,[Website]
            ,[Email]
            ,[Phone]
			,[UniversityId]
            )
        VALUES (
            @FacultyAbbreviation,
			@FacultyName,
			@FacultyWebsite,
			@FacultyEmail,
			@FacultyPhone,
			@UniversityId
            );

		SET @FacultyId = SCOPE_IDENTITY();

        INSERT INTO [dbo].[Departments]
            (
            [Abbreviation]
            ,[Name]
            ,[Website]
            ,[Email]
            ,[Phone]
			,[FacultyId]
            )
        VALUES (
            @DepartmentAbbreviation,
			@DepartmentName,
			@DepartmentWebsite,
			@DepartmentEmail,
			@DepartmentPhone,
			@FacultyId
            );

		SET @DepartmentId = SCOPE_IDENTITY();

        COMMIT TRANSACTION; 
    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0
        BEGIN
            ROLLBACK TRANSACTION MySavePoint; -- rollback to MySavePoint
        END
    END CATCH

END
GO