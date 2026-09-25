SELECT
    tblUserLog.Timestamp,
    tblUserLog.UserComputerName,
    tblUserLog.Activity

FROM
    tblUserLog

WHERE
    (((tblUserLog.Activity)="Login QAPI" Or (tblUserLog.Activity)="Logoff QAPI" Or (tblUserLog.Activity)="Force Logout"));
