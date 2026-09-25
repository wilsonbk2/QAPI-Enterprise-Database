SELECT
    [TrainingClassRoleRequirements P3].*,
    SDTCOURSECOMP.[Last Launch],
    IIf(
IsNull([Last Launch]),
        "Incomplete",
        "Completed"
    )AS CompletionStatus,
    IIf(
        [OKEY]="8",
        "N/A",
        IIf(
            IIf(
IsNull([Last Launch]),
                "Incomplete",
                "Completed"
            )="Completed",
            IIf(
                [Last Launch]>=[DATEASSIGN]
                AND
                [Last Launch]<=[DATEDUE],
                "On-Time",
                IIf(
                    [Last Launch]>[DATEDUE],
                    "Late",
                    IIf(
                        [Last Launch]<[DATEASSIGN],
                        "Early",
                        "Unicorn"
                    )
                )
            ),
            "Incomplete"
        )
    )AS TrainingStatus,
    IIf(
        [OKEY]="8",
        IIf(
IsNull([Last Launch]),
            "Hide",
            "Class Taken"
        ),
        "Requirement"
    )AS [Exception Classes],
    SDTCOURSECOMP.Duplicate

FROM
    [TrainingClassRoleRequirements P3]
    INNER JOIN SDTCOURSECOMP
        ON ([TrainingClassRoleRequirements P3].COURSEID = SDTCOURSECOMP.COURSEID)
        AND
        ([TrainingClassRoleRequirements P3].SDTYEAR = SDTCOURSECOMP.YearComp)
        AND
        ([TrainingClassRoleRequirements P3].GLAIT = SDTCOURSECOMP.[User ID])

WHERE
    ((( IIf(
        [OKEY]="8",
        IIf(
IsNull([Last Launch]),
            "Hide",
            "Class Taken"
        ),
        "Requirement"
    ) )<>"Hide") AND ((SDTCOURSECOMP.Duplicate) Is Null));
