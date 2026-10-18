/*
    Normalise conference.startdate and conference.enddate to padded YYYY-MM-DD.

    conference.startdate is an nchar(10) free text column with no format
    enforcement, and admin/edit.asp used to write three different shapes into
    it depending on the day of the month:

        2026-10-14   padded iso      - written by admin/add.asp
        2027-6-14    unpadded iso    - admin/edit.asp when the day was > 12
        02/06/2027   dd/mm/yyyy      - admin/edit.asp when the day was <= 12

    Anything but the first shape made server/lib.inc parseDbDate() return empty,
    which silently disabled post-registration editing for that conference.

    The admin screens now always write the padded form, so this script only has
    to clean up what is already stored.

    Run STEP 1 and read the result before running STEP 2.
    Through sqlcmd add -I (see the project notes on QUOTED_IDENTIFIER).
*/

USE [imknasim];
GO

-- ---------------------------------------------------------------------------
-- STEP 1 - what would change. Expect newStartdate / newEnddate to be NULL for
-- rows that are already correct.
-- ---------------------------------------------------------------------------
SELECT  c.theindex,
        LEFT(c.thename, 30)                 AS conference,
        '[' + RTRIM(c.startdate) + ']'      AS startdate,
        ISNULL('[' + RTRIM(s.iso) + ']', '(no change)') AS newStartdate,
        '[' + RTRIM(c.enddate) + ']'        AS enddate,
        ISNULL('[' + RTRIM(e.iso) + ']', '(no change)') AS newEnddate
FROM    dbo.conference c
CROSS APPLY (SELECT LTRIM(RTRIM(c.startdate)) AS t) a
CROSS APPLY (SELECT CASE
                WHEN a.t LIKE '[0-9][0-9]/[0-9][0-9]/[0-9][0-9][0-9][0-9]'
                    THEN SUBSTRING(a.t,7,4) + '-' + SUBSTRING(a.t,4,2) + '-' + SUBSTRING(a.t,1,2)
                WHEN a.t LIKE '[0-9][0-9][0-9][0-9]-%-%'
                 AND a.t NOT LIKE '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]'
                    THEN LEFT(a.t,4) + '-'
                       + RIGHT('0' + SUBSTRING(a.t, 6, CHARINDEX('-', a.t, 6) - 6), 2) + '-'
                       + RIGHT('0' + SUBSTRING(a.t, CHARINDEX('-', a.t, 6) + 1, 2), 2)
                END AS iso) s
CROSS APPLY (SELECT LTRIM(RTRIM(c.enddate)) AS t) b
CROSS APPLY (SELECT CASE
                WHEN b.t LIKE '[0-9][0-9]/[0-9][0-9]/[0-9][0-9][0-9][0-9]'
                    THEN SUBSTRING(b.t,7,4) + '-' + SUBSTRING(b.t,4,2) + '-' + SUBSTRING(b.t,1,2)
                WHEN b.t LIKE '[0-9][0-9][0-9][0-9]-%-%'
                 AND b.t NOT LIKE '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]'
                    THEN LEFT(b.t,4) + '-'
                       + RIGHT('0' + SUBSTRING(b.t, 6, CHARINDEX('-', b.t, 6) - 6), 2) + '-'
                       + RIGHT('0' + SUBSTRING(b.t, CHARINDEX('-', b.t, 6) + 1, 2), 2)
                END AS iso) e
WHERE   s.iso IS NOT NULL OR e.iso IS NOT NULL
ORDER BY c.theindex;
GO

-- ---------------------------------------------------------------------------
-- STEP 2 - apply it
-- ---------------------------------------------------------------------------
UPDATE  c
SET     startdate = ISNULL(s.iso, c.startdate),
        enddate   = ISNULL(e.iso, c.enddate)
FROM    dbo.conference c
CROSS APPLY (SELECT LTRIM(RTRIM(c.startdate)) AS t) a
CROSS APPLY (SELECT CASE
                WHEN a.t LIKE '[0-9][0-9]/[0-9][0-9]/[0-9][0-9][0-9][0-9]'
                    THEN SUBSTRING(a.t,7,4) + '-' + SUBSTRING(a.t,4,2) + '-' + SUBSTRING(a.t,1,2)
                WHEN a.t LIKE '[0-9][0-9][0-9][0-9]-%-%'
                 AND a.t NOT LIKE '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]'
                    THEN LEFT(a.t,4) + '-'
                       + RIGHT('0' + SUBSTRING(a.t, 6, CHARINDEX('-', a.t, 6) - 6), 2) + '-'
                       + RIGHT('0' + SUBSTRING(a.t, CHARINDEX('-', a.t, 6) + 1, 2), 2)
                END AS iso) s
CROSS APPLY (SELECT LTRIM(RTRIM(c.enddate)) AS t) b
CROSS APPLY (SELECT CASE
                WHEN b.t LIKE '[0-9][0-9]/[0-9][0-9]/[0-9][0-9][0-9][0-9]'
                    THEN SUBSTRING(b.t,7,4) + '-' + SUBSTRING(b.t,4,2) + '-' + SUBSTRING(b.t,1,2)
                WHEN b.t LIKE '[0-9][0-9][0-9][0-9]-%-%'
                 AND b.t NOT LIKE '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]'
                    THEN LEFT(b.t,4) + '-'
                       + RIGHT('0' + SUBSTRING(b.t, 6, CHARINDEX('-', b.t, 6) - 6), 2) + '-'
                       + RIGHT('0' + SUBSTRING(b.t, CHARINDEX('-', b.t, 6) + 1, 2), 2)
                END AS iso) e
WHERE   s.iso IS NOT NULL OR e.iso IS NOT NULL;
GO

-- ---------------------------------------------------------------------------
-- STEP 3 - verify. Every row should come back as OK.
-- ---------------------------------------------------------------------------
SELECT  theindex,
        LEFT(thename, 30) AS conference,
        RTRIM(startdate)  AS startdate,
        RTRIM(enddate)    AS enddate,
        CASE WHEN RTRIM(startdate) LIKE '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]'
              AND RTRIM(enddate)   LIKE '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]'
             THEN 'OK' ELSE 'STILL WRONG' END AS status
FROM    dbo.conference
ORDER BY theindex;
GO
