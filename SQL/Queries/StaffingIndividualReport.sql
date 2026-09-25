SELECT
    StaffingLog.*,
    EmployeeOnlyALL.*,
    [BUILDING#] & "-" & [UNITMAIN] & "-" & [UNITSUB] AS Location

FROM
    (StaffingLog LEFT JOIN EmployeeOnlyALL ON StaffingLog.EMPID = EmployeeOnlyALL.[EMP ID])
    LEFT JOIN DEQRESOLUTION
        ON StaffingLog.CSCKEY = DEQRESOLUTION.DEQKEY

WHERE
    (((StaffingLog.SDATE) Between Forms!StaffingReports!staffP24 And Forms!StaffingReports!staffP25) And ((EmployeeOnlyALL.[EMP ID]) Like "*" & Forms!StaffingReports!staffP21 & "*") And ((DEQRESOLUTION.DEQKEY) Is Null))

ORDER BY
    StaffingLog.SDATE DESC;
