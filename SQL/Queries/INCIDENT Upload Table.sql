SELECT
    [INCIDENT TRACKER UPLOAD].*,
    Format(
        [Date],
        "Short Date"
    )AS ActualDate,
    Format(
        [Time],
        "General Number"
    )AS ActualTime

FROM
    [INCIDENT TRACKER UPLOAD];
