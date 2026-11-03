/*
    One column on closedlist, so that an invitation is good for exactly one
    registration.

    When somebody completes a registration through the closed list screen,
    their row is marked rather than deleted - the organiser keeps the whole
    invitation list and can see at a glance who has already come through. A
    marked row no longer opens the form, and the form tells that visitor to
    use the edit link instead of saying they were never invited.

    Editing an existing registration does not spend an invitation.

    Safe to run twice: the ALTER only happens when the column is missing.

    Through sqlcmd add -I (see the project notes on QUOTED_IDENTIFIER):
        sqlcmd -S YAN2011 -E -C -I -d imknasim -i sql\add_closedlist_isregistered.sql
*/

USE [imknasim];
GO

IF COL_LENGTH('dbo.closedlist', 'isregistered') IS NULL
    ALTER TABLE dbo.closedlist ADD isregistered bit NOT NULL
        CONSTRAINT DF_closedlist_isregistered DEFAULT 0;
GO

-- Verify. Every existing row starts unused, which is what you want: the list
-- is empty in production today, and nobody has registered through it yet.
SELECT  theindex,
        theconference,
        RTRIM(ISNULL(thename, ''))  AS invitee,
        RTRIM(ISNULL(theemail, '')) AS theemail,
        RTRIM(ISNULL(thephone, '')) AS thephone,
        isregistered
FROM    dbo.closedlist
ORDER BY theconference, theindex;
GO
