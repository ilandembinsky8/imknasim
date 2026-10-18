/*
    Run this after every restore of a backup taken from the production server.

    A backup carries the database users with the SIDs they had on the machine
    they came from. The local server login michal has a different SID, so after
    a restore the user is orphaned and the site cannot connect at all:

        Login failed for user 'michal'.
        Cannot open database "imknasim" requested by the login.

    server/secure.inc starts with On Error Resume Next, so none of that reaches
    the browser - the endpoints just answer {} and the form renders empty. There
    is no error message to go on, which is why this script exists.

    ALTER USER rewrites the stored SID to match the login. It touches neither
    the permissions, nor the password, nor any data.

    Through sqlcmd add -I (see the project notes on QUOTED_IDENTIFIER):
        sqlcmd -S .\SQLEXPRESS -E -C -I -i sql\after_restore_fix_login.sql
*/

USE [imknasim];
GO

ALTER USER [michal] WITH LOGIN = [michal];
GO

-- Verify. Both SIDs must come back identical and the state must be 'linked'.
SELECT  dp.name,
        CONVERT(varchar(64), dp.sid, 1) AS databaseSid,
        CONVERT(varchar(64), sp.sid, 1) AS loginSid,
        CASE WHEN sp.sid IS NULL THEN 'STILL ORPHANED' ELSE 'linked' END AS state
FROM    sys.database_principals dp
LEFT JOIN sys.server_principals sp ON sp.sid = dp.sid
WHERE   dp.name = 'michal';
GO

-- The roles come across with the backup, but check them while we are here.
SELECT  r.name AS membership
FROM    sys.database_role_members m
JOIN    sys.database_principals r ON r.principal_id = m.role_principal_id
JOIN    sys.database_principals u ON u.principal_id = m.member_principal_id
WHERE   u.name = 'michal'
ORDER BY r.name;
GO
