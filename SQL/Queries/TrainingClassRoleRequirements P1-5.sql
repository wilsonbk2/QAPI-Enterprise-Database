SELECT
    IIf(
IsNull([SEPARATION]),
Date(),
        [SEPARATION]
    )AS EmpCurrentDay,
    TrainingAllEmployeeHistory2.*,
    [TrainingClassRoleRequirements P1].COURSEID,
    [TrainingClassRoleRequirements P1].COURSENAME,
    [TrainingClassRoleRequirements P1].COURSEESTABLISH,
    [TrainingClassRoleRequirements P1].REQUIRED,
    [TrainingClassRoleRequirements P1].CLASSTYDESC,
    [TrainingClassRoleRequirements P1].COURSEDESC,
    [TrainingClassRoleRequirements P1].OCCURDESC,
    [TrainingClassRoleRequirements P1].OCCURPERIOD,
    [TrainingClassRoleRequirements P1].TRNTYPDESC,
    [TrainingClassRoleRequirements P1].TRNCATDESC,
    [TrainingClassRoleRequirements P1].TRNCATID2,
    [TrainingClassRoleRequirements P1].OKEY,
    [TrainingClassRoleRequirements P1].SDTYEAR,
    [TrainingClassRoleRequirements P1].SDTDATEYEAR,
    [TrainingClassRoleRequirements P1].INTDATEASSIGN

FROM
    [TrainingClassRoleRequirements P1]
    INNER JOIN TrainingAllEmployeeHistory2
        ON [TrainingClassRoleRequirements P1].[ROLE-CODE] = TrainingAllEmployeeHistory2.[ROLE-CODE]

WHERE
    (((TrainingAllEmployeeHistory2.EMPSTATUS)<>"Position Change") AND ((TrainingAllEmployeeHistory2.[Full Name]) Is Not Null));
