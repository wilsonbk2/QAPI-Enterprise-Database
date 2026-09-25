SELECT
    SDTCOURSECOMP.[Content Title],
    SDTCOURSECOMP.[Last Launch],
    SDTCOURSECOMP.YearComp,
    SDTCOURSECOMP.COURSEID,
    SDTCOURSECOMP.Duplicate,
    SDTCOURSECOMP.DuplicateFind,
    EmployeeOnlyALL.JOBTITLE,
    EmployeeOnlyALL.[BLDG#],
    EmployeeOnlyALL.POD,
    EmployeeOnlyALL.WARD,
    EmployeeOnlyALL.[Supervisor Name],
    EmployeeOnlyALL.[CC-BASE],
    EmployeeOnlyALL.[CC-DESC],
    EmployeeOnlyALL.EMPSTATUS,
    EmployeeOnlyALL.SEPARATION,
    EmployeeOnlyALL.[POSTYP DESC],
    SDTCOURSECOMP.[User ID],
    EmployeeOnlyALL.[Photo Path],
    EmployeeOnlyALL.SupvJobTitle

FROM
    (EmployeeOnlyALL INNER JOIN TrainingComplianceReportsNamesforCOMP ON EmployeeOnlyALL.SSN = TrainingComplianceReportsNamesforCOMP.SSN)
    INNER JOIN SDTCOURSECOMP
        ON TrainingComplianceReportsNamesforCOMP.GLAIT = SDTCOURSECOMP.[User ID]

WHERE
    (((SDTCOURSECOMP.YearComp) Like "*" & Forms!TrainingComplianceReports!EVENT3 & "*") And ((SDTCOURSECOMP.Duplicate) Is Null) And ((SDTCOURSECOMP.[User ID]) Like "*" & Forms!TrainingComplianceReports!EVENT2 & "*"));
