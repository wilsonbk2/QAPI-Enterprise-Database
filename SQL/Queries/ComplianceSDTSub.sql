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
    LEFT JOIN ComplianceSDT
        ON (ComplianceCurrentDateRangeMonths.Month = ComplianceSDT.FirstOfMonth)
        AND
        (ComplianceCurrentDateRangeMonths.Year = ComplianceSDT.Year)

ORDER BY
    ComplianceCurrentDateRangeMonths.Year
    Val(Replace([MonthNo],"-",""));
