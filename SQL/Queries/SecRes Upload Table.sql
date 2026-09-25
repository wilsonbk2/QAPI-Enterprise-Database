SELECT
    [SECRES UPLOAD].*,
    Format(
        [Date/Time Procedure Initated],
        "Short Date"
    )AS ActualDate,
    Format(
        [Date/Time Procedure Initated],
        "General Number"
    )AS ActualTime

FROM
    [SECRES UPLOAD];
