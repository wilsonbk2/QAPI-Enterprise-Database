SELECT
    ComplianceCurrentDateRangeMonths.Year,
    ComplianceCurrentDateRangeMonths.Month,
Val(Replace([MonthNo],"-","")) AS MonNum,
    IIf(
IsNull([FirstOfMonth]),
        "No Data Submission",
        "Successful Submission"
    )AS [Submit Status]

FROM
    ComplianceCurrentDateRangeMonths
    LEFT JOIN ComplianceJRCP
        ON (ComplianceCurrentDateRangeMonths.Month = ComplianceJRCP.FirstOfMonth)
        AND
        (ComplianceCurrentDateRangeMonths.Year = ComplianceJRCP.Year)

ORDER BY
    ComplianceCurrentDateRangeMonths.Year
    Val(Replace([MonthNo],"-",""));
