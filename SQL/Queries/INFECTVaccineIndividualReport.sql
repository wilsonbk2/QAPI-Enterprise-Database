SELECT
    EmployeeOnlyALL.[SUPV#],
    EmployeeOnlyALL.[EMP ID],
    EmployeeOnlyALL.[Full Name],
    EmployeeOnlyALL.JOBTITLE,
    INFECTvaccinelist.VACDESC,
    INFECTvaccinelist.VACDATE,
    INFECTvaccinelist.ExpireDate,
    INFECTvaccinelist.ExpireStatus,
    EmployeeOnlyALL.[CC-BASE],
    EmployeeOnlyALL.[CC-DESC],
    [EmployeeOnlyALL].[BLDG#] & "-" & [EmployeeOnlyALL].[POD] & "-" & [EmployeeOnlyALL].[WARD] AS Location,
    EmployeeOnlyALL.EMPSTATUS,
    EmployeeOnlyALL.SEPARATION,
    EmployeeOnlyALL.[Photo Path],
    EmployeeOnlyALL.[POSTYP DESC],
    EmployeeOnlyALL.[Supervisor Name],
    EmployeeOnlyALL.SupvJobTitle,
    EmployeeOnlyALL.NursingDesignationCode

FROM
    EmployeeOnlyALL
    LEFT JOIN INFECTvaccinelist
        ON EmployeeOnlyALL.[EMP ID] = INFECTvaccinelist.EID

WHERE
    (((EmployeeOnlyALL.[SUPV#]) Is Not Null) And ((EmployeeOnlyALL.[EMP ID]) Like "*" & Forms!INFECTHandWashReports!EVENT2 & "*"));
