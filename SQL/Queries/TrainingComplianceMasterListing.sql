SELECT
    TrainingRequiredClassesbyRoleYear.SDTDATEYEAR,
    TrainingRequiredClassesbyRoleYear.[SUPV#],
    TrainingRequiredClassesbyRoleYear.[EMP ID],
    TrainingRequiredClassesbyRoleYear.*,
    TrainingEmpClassCompletion.[Content Title],
    TrainingEmpClassCompletion.[User ID],
    TrainingEmpClassCompletion.[Last Launch],
    TrainingEmpClassCompletion.[Total Launches],
    TrainingEmpClassCompletion.YearComp,
Year([COURSEESTABLISH]) AS EstabYear,
    IIf(
        [SDTYEAR]= IIf(
            [OCCURDESC]="One-Time Class",
Year([COURSEESTABLISH]),
Year([SDTDATEYEAR])
        ),
        "Year of Class",
        "Not Year of Class"
    )AS ClassNeedIndicator,
    DateAdd(
        "q",
        Format(
            [COURSEESTABLISH],
            "q"
        ),
        DateSerial(
Year([SDTDATEYEAR]),
            1,
            1
        )
    )-1 AS QuarterComp,
    DateSerial(
Year([SDTDATEYEAR]),
Month([COURSEESTABLISH]),
Day([COURSEESTABLISH])
    )+90 AS AnnualComp,
    IIf(
IsNull([YearComp]),
        "Incomplete",
        "Complete"
    )AS RecordStatus,
    IIf(
        [OCCURPERIOD]="A",
        [AnnualComp],
        [QuarterComp]
    )AS EndDate,
Date() AS CurDate,
    IIf(
        [RecordStatus]="Complete",
        IIf(
            [Last Launch]>=[COURSEESTABLISH]
            AND
            [Last Launch]<=[EndDate],
            "Complete",
            IIf(
                [Last Launch]<[COURSEESTABLISH],
                "Assignment Early",
                "Assignment Late"
            )
        ),
        IIf(
            [CurDate]>[EndDate],
            "Deficient",
            IIf(
                [CurDate]<[SequentialDate],
                "Class Not Started",
                "Pending Completion"
            )
        )
    )AS CompletionStatus,
    DateSerial(
Year([SDTDATEYEAR]),
Month([COURSEESTABLISH]),
Day([COURSEESTABLISH])
    )AS SequentialDate,
    IIf(
        [OCCURDESC]="One-Time Class",
        [COURSEESTABLISH],
        [SequentialDate]
    )AS FinalizedStartDate,
Year([FinalizedStartDate]) AS FinalizedStartYear,
    IIf(
IsNull([SEPARATION]),
Date(),
        [SEPARATION]
    )AS [Active Date],
    IIf(
        [EMPSTATUS]="Current",
        IIf(
            [ESH-HIRE-DATE]< IIf(
                [OCCURPERIOD]="A",
                DateSerial(
Year([SDTDATEYEAR]),
Month([COURSEESTABLISH]),
Day([COURSEESTABLISH])
                )+90,
                DateAdd(
                    "q",
                    Format(
                        [COURSEESTABLISH],
                        "q"
                    ),
                    DateSerial(
Year([SDTDATEYEAR]),
                        1,
                        1
                    )
                )-1
            ),
            "Non-Exempt",
            "Exempt"
        ),
        IIf(
            [SEPARATION]< IIf(
                [OCCURDESC]="One-Time Class",
                [COURSEESTABLISH],
                DateSerial(
Year([SDTDATEYEAR]),
Month([COURSEESTABLISH]),
Day([COURSEESTABLISH])
                )
            ),
            "Exempt",
            "Non-Exempt"
        )
    )AS Exemption,
    SERVICEBREAK.[EXIT-DATE],
    SERVICEBREAK.[POS-RESTART-DATE],
    IIf(
        [EMPSTATUS]="Current",
        IIf(
            [POS-RESTART-DATE]<= IIf(
                [OCCURPERIOD]="A",
                DateSerial(
Year([SDTDATEYEAR]),
Month([COURSEESTABLISH]),
Day([COURSEESTABLISH])
                )+90,
                DateAdd(
                    "q",
                    Format(
                        [COURSEESTABLISH],
                        "q"
                    ),
                    DateSerial(
Year([SDTDATEYEAR]),
                        1,
                        1
                    )
                )-1
            ),
            "Required1",
            IIf(
                [EXIT-DATE]> IIf(
                    [OCCURPERIOD]="A",
                    DateSerial(
Year([SDTDATEYEAR]),
Month([COURSEESTABLISH]),
Day([COURSEESTABLISH])
                    )+90,
                    DateAdd(
                        "q",
                        Format(
                            [COURSEESTABLISH],
                            "q"
                        ),
                        DateSerial(
Year([SDTDATEYEAR]),
                            1,
                            1
                        )
                    )-1
                ),
                "Not Required Service Break",
                IIf(
                    [EXIT-DATE]> IIf(
                        [OCCURDESC]="One-Time Class",
                        [COURSEESTABLISH],
                        DateSerial(
Year([SDTDATEYEAR]),
Month([COURSEESTABLISH]),
Day([COURSEESTABLISH])
                        )
                    )
                    AND
                    [EXIT-DATE]<= IIf(
                        [OCCURPERIOD]="A",
                        DateSerial(
Year([SDTDATEYEAR]),
Month([COURSEESTABLISH]),
Day([COURSEESTABLISH])
                        )+90,
                        DateAdd(
                            "q",
                            Format(
                                [COURSEESTABLISH],
                                "q"
                            ),
                            DateSerial(
Year([SDTDATEYEAR]),
                                1,
                                1
                            )
                        )-1
                    ),
                    "Required2",
                    IIf(
IsNull([EXIT-DATE]),
                        "No Service Break Recorded",
                        "Not Required Outlier"
                    )
                )
            )
        ),
        "Currently Separated "
    )AS ServiceBreak1,
    IIf(
        [POS-RESTART-DATE]<= IIf(
            [OCCURPERIOD]="A",
            DateSerial(
Year([SDTDATEYEAR]),
Month([COURSEESTABLISH]),
Day([COURSEESTABLISH])
            )+90,
            DateAdd(
                "q",
                Format(
                    [COURSEESTABLISH],
                    "q"
                ),
                DateSerial(
Year([SDTDATEYEAR]),
                    1,
                    1
                )
            )-1
        ),
        [C2],
        [C3]
    )AS C1,
    IIf(
        [POS-RESTART-DATE]<[TrainingRequiredClassesbyRoleYear].[POS-START-DATE],
        "Filter Out Previous Service Break",
        "Required1"
    )AS C2,
    IIf(
        IIf(
            [OCCURPERIOD]="A",
            DateSerial(
Year([SDTDATEYEAR]),
Month([COURSEESTABLISH]),
Day([COURSEESTABLISH])
            )+90,
            DateAdd(
                "q",
                Format(
                    [COURSEESTABLISH],
                    "q"
                ),
                DateSerial(
Year([SDTDATEYEAR]),
                    1,
                    1
                )
            )-1
        )>=[EXIT-DATE]
        AND
        IIf(
            [OCCURPERIOD]="A",
            DateSerial(
Year([SDTDATEYEAR]),
Month([COURSEESTABLISH]),
Day([COURSEESTABLISH])
            )+90,
            DateAdd(
                "q",
                Format(
                    [COURSEESTABLISH],
                    "q"
                ),
                DateSerial(
Year([SDTDATEYEAR]),
                    1,
                    1
                )
            )-1
        )<=[POS-RESTART-DATE],
        "Required3",
        [C4]
    )AS C3,
    IIf(
        [POS-RESTART-DATE]> IIf(
            [OCCURPERIOD]="A",
            DateSerial(
Year([SDTDATEYEAR]),
Month([COURSEESTABLISH]),
Day([COURSEESTABLISH])
            )+90,
            DateAdd(
                "q",
                Format(
                    [COURSEESTABLISH],
                    "q"
                ),
                DateSerial(
Year([SDTDATEYEAR]),
                    1,
                    1
                )
            )-1
        ),
        "Not required Restarted After Class End Date",
        "Required4"
    )AS C4,
    IIf(
        [ServiceBreak1]="Required1",
        IIf(
            [POS-RESTART-DATE]<=[EndDate],
            IIf(
                [POS-RESTART-DATE]<[TrainingRequiredClassesbyRoleYear].[POS-START-DATE],
                "Filter Out Previous Service Break",
                "Required1"
            ),
            IIf(
                [EndDate]>=[EXIT-DATE]
                AND
                [EndDate]<=[POS-RESTART-DATE],
                "Required3",
                IIf(
                    [POS-RESTART-DATE]>[EndDate],
                    "Not Required Restarted After Class End Date",
                    "Required4"
                )
            )
        ),
        [ServiceBreak1]
    )AS ServiceBreakFinalOLD,
    IIf(
        [ServiceBreak1]="Required",
        [C1],
        [ServiceBreak1]
    )AS ServiceBreakFinal

FROM
    SERVICEBREAK
    RIGHT JOIN (TrainingRequiredClassesbyRoleYear LEFT JOIN TrainingEmpClassCompletion
        ON (TrainingRequiredClassesbyRoleYear.GLAIT = TrainingEmpClassCompletion.[User ID])
        AND
        (TrainingRequiredClassesbyRoleYear.COURSENAME = TrainingEmpClassCompletion.[Content Title])
        AND
        (TrainingRequiredClassesbyRoleYear.SDTYEAR = TrainingEmpClassCompletion.YearComp)) ON SERVICEBREAK.EMPID = TrainingRequiredClassesbyRoleYear.[EMP ID]

WHERE
    (
        (
            IIf(
                [SDTYEAR] = IIf(
                    [OCCURDESC] = "One-Time Class",
                    Year([COURSEESTABLISH]),
                    Year([SDTDATEYEAR])
                ),
                "Year of Class",
                "Not Year of Class"
            )
        ) = "Year of Class"

        AND

        (
            IIf(
                [EMPSTATUS] = "Current",
                IIf(
                    [ESH-HIRE-DATE] < IIf(
                        [OCCURPERIOD] = "A",
                        DateSerial(
                            Year([SDTDATEYEAR]),
                            Month([COURSEESTABLISH]),
                            Day([COURSEESTABLISH])
                        ) + 90,
                        DateAdd(
                            "q",
                            Format([COURSEESTABLISH], "q"),
                            DateSerial(
                                Year([SDTDATEYEAR]),
                                1,
                                1
                            )
                        ) - 1
                    ),
                    "Non-Exempt",
                    "Exempt"
                ),
                IIf(
                    [SEPARATION] < IIf(
                        [OCCURDESC] = "One-Time Class",
                        [COURSEESTABLISH],
                        DateSerial(
                            Year([SDTDATEYEAR]),
                            Month([COURSEESTABLISH]),
                            Day([COURSEESTABLISH])
                        )
                    ),
                    "Exempt",
                    "Non-Exempt"
                )
            )
        ) = "Non-Exempt"

        AND

        (
            IIf(
                [EMPSTATUS] = "Current",
                IIf(
                    [POS-RESTART-DATE] <= IIf(
                        [OCCURPERIOD] = "A",
                        DateSerial(
                            Year([SDTDATEYEAR]),
                            Month([COURSEESTABLISH]),
                            Day([COURSEESTABLISH])
                        ) + 90,
                        DateAdd(
                            "q",
                            Format([COURSEESTABLISH], "q"),
                            DateSerial(
                                Year([SDTDATEYEAR]),
                                1,
                                1
                            )
                        ) - 1
                    ),
                    "Required1",
                    IIf(
                        [EXIT-DATE] > IIf(
                            [OCCURPERIOD] = "A",
                            DateSerial(
                                Year([SDTDATEYEAR]),
                                Month([COURSEESTABLISH]),
                                Day([COURSEESTABLISH])
                            ) + 90,
                            DateAdd(
                                "q",
                                Format([COURSEESTABLISH], "q"),
                                DateSerial(
                                    Year([SDTDATEYEAR]),
                                    1,
                                    1
                                )
                            ) - 1
                        ),
                        "Not Required Service Break",
                        IIf(
                            [EXIT-DATE] > IIf(
                                [OCCURDESC] = "One-Time Class",
                                [COURSEESTABLISH],
                                DateSerial(
                                    Year([SDTDATEYEAR]),
                                    Month([COURSEESTABLISH]),
                                    Day([COURSEESTABLISH])
                                )
                            )
                            AND
                            [EXIT-DATE] <= IIf(
                                [OCCURPERIOD] = "A",
                                DateSerial(
                                    Year([SDTDATEYEAR]),
                                    Month([COURSEESTABLISH]),
                                    Day([COURSEESTABLISH])
                                ) + 90,
                                DateAdd(
                                    "q",
                                    Format([COURSEESTABLISH], "q"),
                                    DateSerial(
                                        Year([SDTDATEYEAR]),
                                        1,
                                        1
                                    )
                                ) - 1
                            ),
                            "Required2",
                            IIf(
                                IsNull([EXIT-DATE]),
                                "No Service Break Recorded",
                                "Not Required Outlier"
                            )
                        )
                    )
                ),
                "Currently Separated "
            )
        ) = "Required1"

        OR

        (
            IIf(
                [EMPSTATUS] = "Current",
                IIf(
                    [POS-RESTART-DATE] <= IIf(
                        [OCCURPERIOD] = "A",
                        DateSerial(
                            Year([SDTDATEYEAR]),
                            Month([COURSEESTABLISH]),
                            Day([COURSEESTABLISH])
                        ) + 90,
                        DateAdd(
                            "q",
                            Format([COURSEESTABLISH], "q"),
                            DateSerial(
                                Year([SDTDATEYEAR]),
                                1,
                                1
                            )
                        ) - 1
                    ),
                    "Required1",
                    IIf(
                        [EXIT-DATE] > IIf(
                            [OCCURPERIOD] = "A",
                            DateSerial(
                                Year([SDTDATEYEAR]),
                                Month([COURSEESTABLISH]),
                                Day([COURSEESTABLISH])
                            ) + 90,
                            DateAdd(
                                "q",
                                Format([COURSEESTABLISH], "q"),
                                DateSerial(
                                    Year([SDTDATEYEAR]),
                                    1,
                                    1
                                )
                            ) - 1
                        ),
                        "Not Required Service Break",
                        IIf(
                            [EXIT-DATE] > IIf(
                                [OCCURDESC] = "One-Time Class",
                                [COURSEESTABLISH],
                                DateSerial(
                                    Year([SDTDATEYEAR]),
                                    Month([COURSEESTABLISH]),
                                    Day([COURSEESTABLISH])
                                )
                            )
                            AND
                            [EXIT-DATE] <= IIf(
                                [OCCURPERIOD] = "A",
                                DateSerial(
                                    Year([SDTDATEYEAR]),
                                    Month([COURSEESTABLISH]),
                                    Day([COURSEESTABLISH])
                                ) + 90,
                                DateAdd(
                                    "q",
                                    Format([COURSEESTABLISH], "q"),
                                    DateSerial(
                                        Year([SDTDATEYEAR]),
                                        1,
                                        1
                                    )
                                ) - 1
                            ),
                            "Required2",
                            IIf(
                                IsNull([EXIT-DATE]),
                                "No Service Break Recorded",
                                "Not Required Outlier"
                            )
                        )
                    )
                ),
                "Currently Separated "
            )
        ) = "Currently Separated"

        OR

        (
            IIf(
                [EMPSTATUS] = "Current",
                IIf(
                    [POS-RESTART-DATE] <= IIf(
                        [OCCURPERIOD] = "A",
                        DateSerial(
                            Year([SDTDATEYEAR]),
                            Month([COURSEESTABLISH]),
                            Day([COURSEESTABLISH])
                        ) + 90,
                        DateAdd(
                            "q",
                            Format([COURSEESTABLISH], "q"),
                            DateSerial(
                                Year([SDTDATEYEAR]),
                                1,
                                1
                            )
                        ) - 1
                    ),
                    "Required1",
                    IIf(
                        [EXIT-DATE] > IIf(
                            [OCCURPERIOD] = "A",
                            DateSerial(
                                Year([SDTDATEYEAR]),
                                Month([COURSEESTABLISH]),
                                Day([COURSEESTABLISH])
                            ) + 90,
                            DateAdd(
                                "q",
                                Format([COURSEESTABLISH], "q"),
                                DateSerial(
                                    Year([SDTDATEYEAR]),
                                    1,
                                    1
                                )
                            ) - 1
                        ),
                        "Not Required Service Break",
                        IIf(
                            [EXIT-DATE] > IIf(
                                [OCCURDESC] = "One-Time Class",
                                [COURSEESTABLISH],
                                DateSerial(
                                    Year([SDTDATEYEAR]),
                                    Month([COURSEESTABLISH]),
                                    Day([COURSEESTABLISH])
                                )
                            )
                            AND
                            [EXIT-DATE] <= IIf(
                                [OCCURPERIOD] = "A",
                                DateSerial(
                                    Year([SDTDATEYEAR]),
                                    Month([COURSEESTABLISH]),
                                    Day([COURSEESTABLISH])
                                ) + 90,
                                DateAdd(
                                    "q",
                                    Format([COURSEESTABLISH], "q"),
                                    DateSerial(
                                        Year([SDTDATEYEAR]),
                                        1,
                                        1
                                    )
                                ) - 1
                            ),
                            "Required2",
                            IIf(
                                IsNull([EXIT-DATE]),
                                "No Service Break Recorded",
                                "Not Required Outlier"
                            )
                        )
                    )
                ),
                "Currently Separated "
            )
        ) = "No Service Break Recorded"

        OR

        (
            IIf(
                [EMPSTATUS] = "Current",
                IIf(
                    [POS-RESTART-DATE] <= IIf(
                        [OCCURPERIOD] = "A",
                        DateSerial(
                            Year([SDTDATEYEAR]),
                            Month([COURSEESTABLISH]),
                            Day([COURSEESTABLISH])
                        ) + 90,
                        DateAdd(
                            "q",
                            Format([COURSEESTABLISH], "q"),
                            DateSerial(
                                Year([SDTDATEYEAR]),
                                1,
                                1
                            )
                        ) - 1
                    ),
                    "Required1",
                    IIf(
                        [EXIT-DATE] > IIf(
                            [OCCURPERIOD] = "A",
                            DateSerial(
                                Year([SDTDATEYEAR]),
                                Month([COURSEESTABLISH]),
                                Day([COURSEESTABLISH])
                            ) + 90,
                            DateAdd(
                                "q",
                                Format([COURSEESTABLISH], "q"),
                                DateSerial(
                                    Year([SDTDATEYEAR]),
                                    1,
                                    1
                                )
                            ) - 1
                        ),
                        "Not Required Service Break",
                        IIf(
                            [EXIT-DATE] > IIf(
                                [OCCURDESC] = "One-Time Class",
                                [COURSEESTABLISH],
                                DateSerial(
                                    Year([SDTDATEYEAR]),
                                    Month([COURSEESTABLISH]),
                                    Day([COURSEESTABLISH])
                                )
                            )
                            AND
                            [EXIT-DATE] <= IIf(
                                [OCCURPERIOD] = "A",
                                DateSerial(
                                    Year([SDTDATEYEAR]),
                                    Month([COURSEESTABLISH]),
                                    Day([COURSEESTABLISH])
                                ) + 90,
                                DateAdd(
                                    "q",
                                    Format([COURSEESTABLISH], "q"),
                                    DateSerial(
                                        Year([SDTDATEYEAR]),
                                        1,
                                        1
                                    )
                                ) - 1
                            ),
                            "Required2",
                            IIf(
                                IsNull([EXIT-DATE]),
                                "No Service Break Recorded",
                                "Not Required Outlier"
                            )
                        )
                    )
                ),
                "Currently Separated "
            )
        ) = "Required2"

        OR

        (
            IIf(
                [EMPSTATUS] = "Current",
                IIf(
                    [POS-RESTART-DATE] <= IIf(
                        [OCCURPERIOD] = "A",
                        DateSerial(
                            Year([SDTDATEYEAR]),
                            Month([COURSEESTABLISH]),
                            Day([COURSEESTABLISH])
                        ) + 90,
                        DateAdd(
                            "q",
                            Format([COURSEESTABLISH], "q"),
                            DateSerial(
                                Year([SDTDATEYEAR]),
                                1,
                                1
                            )
                        ) - 1
                    ),
                    "Required1",
                    IIf(
                        [EXIT-DATE] > IIf(
                            [OCCURPERIOD] = "A",
                            DateSerial(
                                Year([SDTDATEYEAR]),
                                Month([COURSEESTABLISH]),
                                Day([COURSEESTABLISH])
                            ) + 90,
                            DateAdd(
                                "q",
                                Format([COURSEESTABLISH], "q"),
                                DateSerial(
                                    Year([SDTDATEYEAR]),
                                    1,
                                    1
                                )
                            ) - 1
                        ),
                        "Not Required Service Break",
                        IIf(
                            [EXIT-DATE] > IIf(
                                [OCCURDESC] = "One-Time Class",
                                [COURSEESTABLISH],
                                DateSerial(
                                    Year([SDTDATEYEAR]),
                                    Month([COURSEESTABLISH]),
                                    Day([COURSEESTABLISH])
                                )
                            )
                            AND
                            [EXIT-DATE] <= IIf(
                                [OCCURPERIOD] = "A",
                                DateSerial(
                                    Year([SDTDATEYEAR]),
                                    Month([COURSEESTABLISH]),
                                    Day([COURSEESTABLISH])
                                ) + 90,
                                DateAdd(
                                    "q",
                                    Format([COURSEESTABLISH], "q"),
                                    DateSerial(
                                        Year([SDTDATEYEAR]),
                                        1,
                                        1
                                    )
                                ) - 1
                            ),
                            "Required2",
                            IIf(
                                IsNull([EXIT-DATE]),
                                "No Service Break Recorded",
                                "Not Required Outlier"
                            )
                        )
                    )
                ),
                "Currently Separated "
            )
        ) = "Required3"

        OR

        (
            IIf(
                [EMPSTATUS] = "Current",
                IIf(
                    [POS-RESTART-DATE] <= IIf(
                        [OCCURPERIOD] = "A",
                        DateSerial(
                            Year([SDTDATEYEAR]),
                            Month([COURSEESTABLISH]),
                            Day([COURSEESTABLISH])
                        ) + 90,
                        DateAdd(
                            "q",
                            Format([COURSEESTABLISH], "q"),
                            DateSerial(
                                Year([SDTDATEYEAR]),
                                1,
                                1
                            )
                        ) - 1
                    ),
                    "Required1",
                    IIf(
                        [EXIT-DATE] > IIf(
                            [OCCURPERIOD] = "A",
                            DateSerial(
                                Year([SDTDATEYEAR]),
                                Month([COURSEESTABLISH]),
                                Day([COURSEESTABLISH])
                            ) + 90,
                            DateAdd(
                                "q",
                                Format([COURSEESTABLISH], "q"),
                                DateSerial(
                                    Year([SDTDATEYEAR]),
                                    1,
                                    1
                                )
                            ) - 1
                        ),
                        "Not Required Service Break",
                        IIf(
                            [EXIT-DATE] > IIf(
                                [OCCURDESC] = "One-Time Class",
                                [COURSEESTABLISH],
                                DateSerial(
                                    Year([SDTDATEYEAR]),
                                    Month([COURSEESTABLISH]),
                                    Day([COURSEESTABLISH])
                                )
                            )
                            AND
                            [EXIT-DATE] <= IIf(
                                [OCCURPERIOD] = "A",
                                DateSerial(
                                    Year([SDTDATEYEAR]),
                                    Month([COURSEESTABLISH]),
                                    Day([COURSEESTABLISH])
                                ) + 90,
                                DateAdd(
                                    "q",
                                    Format([COURSEESTABLISH], "q"),
                                    DateSerial(
                                        Year([SDTDATEYEAR]),
                                        1,
                                        1
                                    )
                                ) - 1
                            ),
                            "Required2",
                            IIf(
                                IsNull([EXIT-DATE]),
                                "No Service Break Recorded",
                                "Not Required Outlier"
                            )
                        )
                    )
                ),
                "Currently Separated "
            )
        ) = "Required4"
    )
)

ORDER BY
    TrainingRequiredClassesbyRoleYear.SDTDATEYEAR,
    TrainingRequiredClassesbyRoleYear.[SUPV#],
    TrainingRequiredClassesbyRoleYear.[EMP ID];

