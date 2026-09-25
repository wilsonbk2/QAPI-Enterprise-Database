SELECT
    [TrainingClassRoleRequirements P4 IP].EMPID,
    [TrainingClassRoleRequirements P4 IP].SDTYEAR,
Count([TrainingClassRoleRequirements P4 IP].COURSEID) AS [Total Of COURSEID]

FROM
    [TrainingClassRoleRequirements P4 IP]

GROUP BY
    [TrainingClassRoleRequirements P4 IP].EMPID
    [TrainingClassRoleRequirements P4 IP].SDTYEAR PIVOT [TrainingClassRoleRequirements P4 IP].CompletionStatus In ("Completed","Incomplete");
