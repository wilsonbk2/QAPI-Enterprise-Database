SELECT
    IIf(
IsNull([Last Launch]),
        "Incomplete",
        "Completed"
    )AS CompletionStatus,
    [TrainingClassRoleRequirements P3].SDTYEAR,
    [TrainingClassRoleRequirements P3].EMPID,
    [TrainingClassRoleRequirements P3].TRNCATDESC,
    [TrainingClassRoleRequirements P3].COURSEID,
    IIf(
        [OKEY]="8",
        IIf(
IsNull([Last Launch]),
            "Hide",
            "Class Taken"
        ),
        "Requirement"
    )AS [Exception Classes]

FROM
    [TrainingClassRoleRequirements P3]
    LEFT JOIN SDTCOURSECOMP
        ON ([TrainingClassRoleRequirements P3].GLAIT = SDTCOURSECOMP.[User ID])
        AND
        ([TrainingClassRoleRequirements P3].SDTYEAR = SDTCOURSECOMP.YearComp)
        AND
        ([TrainingClassRoleRequirements P3].COURSENAME = SDTCOURSECOMP.[Content Title])

WHERE
    ((([TrainingClassRoleRequirements P3].SDTYEAR)<= Year(Date()) And ([TrainingClassRoleRequirements P3].SDTYEAR)>Year(Date())-"2") AND ((IIf([OKEY]="8",IIf(IsNull([Last Launch]),"Hide","Class Taken"),"Requirement"))<>"Hide"));
