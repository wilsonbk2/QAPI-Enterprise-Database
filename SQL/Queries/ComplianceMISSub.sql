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
    LEFT JOIN ComplianceMIS
        ON (ComplianceCurrentDateRangeMonths.Month = ComplianceMIS.FirstOfMonth)
        AND
        (ComplianceCurrentDateRangeMonths.Year = ComplianceMIS.Year)

ORDER BY
    ComplianceCurrentDateRangeMonths.Year
    Val(Replace([MonthNo],"-",""));
