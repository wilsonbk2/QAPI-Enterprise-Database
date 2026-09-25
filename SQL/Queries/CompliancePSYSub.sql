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
    LEFT JOIN CompliancePSY
        ON (ComplianceCurrentDateRangeMonths.Year = CompliancePSY.Year)
        AND
        (ComplianceCurrentDateRangeMonths.Month = CompliancePSY.FirstOfMonth)

ORDER BY
    ComplianceCurrentDateRangeMonths.Year
    Val(Replace([MonthNo],"-",""));
