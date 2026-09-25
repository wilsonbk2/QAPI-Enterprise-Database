SELECT
    [TrainingClassRoleRequirements P4 IP].EMPID,
    [TrainingClassRoleRequirements P4 IP].SDTYEAR,
    [EMPID] & "" & [SDTYEAR] AS EMPYEAR

FROM
    [TrainingClassRoleRequirements P4 IP]

GROUP BY
    [TrainingClassRoleRequirements P4 IP].EMPID
    [TrainingClassRoleRequirements P4 IP].SDTYEAR
    [EMPID] & "" & [SDTYEAR];
