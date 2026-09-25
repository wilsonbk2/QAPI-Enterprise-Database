SELECT
    SDTCOURSE.COURSEID,
    [Forms]![TrainingAddQuarterlyAMTDates]![CB1] AS QYear,
    [Forms]![TrainingAddQuarterlyAMTDates]![CB2] AS QQtr

FROM
    SDTCOURSE

WHERE
    (((SDTCOURSE.IsSelected)=Yes));
