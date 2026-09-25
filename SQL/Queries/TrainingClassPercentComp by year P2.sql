SELECT
    [TrainingClassPercentComp by year P1].EMPID,
    [TrainingClassPercentComp by year P1].SDTYEAR,
    Format(
        IIf(
IsNull([Completed]),
            0,
            [Completed]/[Total Of COURSEID]
        ),
        "Fixed"
    )AS [Percent Complete]

FROM
    [TrainingClassPercentComp by year P1];
