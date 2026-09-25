SELECT
    [TrainingClassPercentComp P1].EMPID,
    [TrainingClassPercentComp P1].TRNCATDESC,
    [TrainingClassPercentComp P1].SDTYEAR,
    Format(
        IIf(
IsNull([Completed]),
            0,
            [Completed]/[Total Of COURSEID]
        ),
        "Fixed"
    )AS [Percent Complete]

FROM
    [TrainingClassPercentComp P1]

WHERE
    ((([TrainingClassPercentComp P1].SDTYEAR)<=Year(Date())));
