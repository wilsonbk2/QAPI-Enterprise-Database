SELECT
    [TrainingClassRoleRequirements P1-5].*,
    TrainingAddQuarterlyCurrentlyAssignedDates.QTRID,
    IIf(
IsNull([QTRID])
        AND
        [OKEY]="9",
        "Remove",
        "Keep"
    )AS QTRASSIGN,
    IIf(
        [OKEY]="9",
        IIf(
IsNull([QID]),
            Null,
            DateSerial(
                [QYEAR],
                [QMONTHASSIGN],
                [QDAYASSIGN]
            )
        ),
        IIf(
            [OKEY]="1"
            OR
            [OKEY]="2",
            [INTDATEASSIGN],
            DateSerial(
                [SDTYEAR],
Month([A-ESH-HIRE-DATE]),
Day([A-ESH-HIRE-DATE])
            )
        )
    )AS DATEASSIGN,
    IIf(
        [OKEY]="9",
        IIf(
IsNull([QID]),
            Null,
            DateSerial(
                [QYEAR],
                [QMONTHDUE],
                [QDAYDUE]
            )
        ),
        IIf(
            [OKEY]="1"
            OR
            [OKEY]="2",
            [INTDATEASSIGN],
            DateSerial(
                [SDTYEAR],
Month([A-ESH-HIRE-DATE]),
Day([A-ESH-HIRE-DATE])
            )
        )+Format([OCCURPERIOD],"Standard")
    )AS DATEDUE,
    IIf(
        [TRNTYPDESC]="ORIENTATION",
        IIf(
            [SDTYEAR]=Year([A-ESH-HIRE-DATE]),
            "Required",
            "Exempt"
        ),
        IIf(
            [TRNTYPDESC]="IN-SERVICE",
            IIf(
                [SDTYEAR]>Year([A-ESH-HIRE-DATE]),
                "Required",
                "Exempt"
            ),
            "Not AMT or Orientation"
        )
    )AS NEOexemption,
    IIf(
        IIf(
IsNull([SEPARATION]),
Date(),
            [SEPARATION]
        )<IIf([OKEY]="9",IIf(IsNull([QID]),Null,DateSerial([QYEAR],[QMONTHASSIGN],[QDAYASSIGN])),IIf([OKEY]="1" Or [OKEY]="2",[INTDATEASSIGN],DateSerial([SDTYEAR],Month([A-ESH-HIRE-DATE]),Day([A-ESH-HIRE-DATE]))))
        OR
        [POS-START-DATE]> IIf(
            [OKEY]="9",
            IIf(
IsNull([QID]),
                Null,
                DateSerial(
                    [QYEAR],
                    [QMONTHDUE],
                    [QDAYDUE]
                )
            ),
            IIf(
                [OKEY]="1"
                OR
                [OKEY]="2",
                [INTDATEASSIGN],
                DateSerial(
                    [SDTYEAR],
Month([A-ESH-HIRE-DATE]),
Day([A-ESH-HIRE-DATE])
                )
            )+Format([OCCURPERIOD],"Standard")
        ),
        IIf(
            [EMPSTATUS]="Current"
            AND
            [SDTYEAR]>=Year(Date()),
            "Future Class - Required",
            "Exempt"
        ),
        "Required"
    )AS SepExempt,
    IIf(
        [DATEASSIGN]>Date(),
        "Future Class",
        IIf(
            [SDTYEAR]<Year(Date()),
            "Past Class",
            "Current Class"
        )
    )AS CurrentorFuturePast

FROM
    [TrainingClassRoleRequirements P1-5]
    LEFT JOIN TrainingAddQuarterlyCurrentlyAssignedDates
        ON ([TrainingClassRoleRequirements P1-5].COURSEID = TrainingAddQuarterlyCurrentlyAssignedDates.COURSEID)
        AND
        ([TrainingClassRoleRequirements P1-5].SDTYEAR = TrainingAddQuarterlyCurrentlyAssignedDates.QYEAR)

WHERE
    ((( IIf(
IsNull([QTRID])
        AND
        [OKEY]="9",
        "Remove",
        "Keep"
    ) )="Keep") AND ((IIf([TRNTYPDESC]="ORIENTATION",IIf([SDTYEAR]=Year([A-ESH-HIRE-DATE]),"Required","Exempt"),IIf([TRNTYPDESC]="IN-SERVICE",IIf([SDTYEAR]>Year([A-ESH-HIRE-DATE]),"Required","Exempt"),"Not AMT or Orientation")))<>"Exempt") AND ((IIf(IIf(IsNull([SEPARATION]),Date(),[SEPARATION])<IIf([OKEY]="9",IIf(IsNull([QID]),Null,DateSerial([QYEAR],[QMONTHASSIGN],[QDAYASSIGN])),IIf([OKEY]="1" Or [OKEY]="2",[INTDATEASSIGN],DateSerial([SDTYEAR],Month([A-ESH-HIRE-DATE]),Day([A-ESH-HIRE-DATE])))) Or [POS-START-DATE]>IIf([OKEY]="9",IIf(IsNull([QID]),Null,DateSerial([QYEAR],[QMONTHDUE],[QDAYDUE])),IIf([OKEY]="1" Or [OKEY]="2",[INTDATEASSIGN],DateSerial([SDTYEAR],Month([A-ESH-HIRE-DATE]),Day([A-ESH-HIRE-DATE])))+Format([OCCURPERIOD],"Standard")),IIf([EMPSTATUS]="Current" And [SDTYEAR]>=Year(Date()),"Future Class - Required","Exempt"),"Required"))<>"Exempt"));
