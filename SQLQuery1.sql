--CREATE TABLE [dbo].[Person](
--	[Id] [int] IDENTITY(1,1) NOT NULL,
--	[FirsName] [varchar](100) NULL,
--	[LastName] [varchar](100) NULL,
--	[Type] [varchar](100) NULL,
--	[AttendeeStatus] [varchar](100) NULL,
--	[YearsExperienceCanada] [varchar](100) NULL,
--	[ApplyingTo] [varchar](100) NULL,
--	[OrganizationMember] [varchar](100) NULL,
--	[LegallyWorkCanada] [varchar](100) NULL,
--	[AgeGroup] [varchar](100) NULL,
--	[MentorshipBefore] [varchar](100) NULL,
--	[IndustryExperience] [varchar](400) NULL,
--	[ProfessionalInterest] [varchar](400) NULL,
--   Assigned varchar (100) NULL
-- CONSTRAINT [PK_Person] PRIMARY KEY CLUSTERED 
--(
--	[Id] ASC
--)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
--) ON [PRIMARY]

--drop table [Person]

--select * from Person
--delete from Person

select * from Person 
where Type = 'Mentee Registration Deposit'
and MentorshipBefore = 'No'
order by type

