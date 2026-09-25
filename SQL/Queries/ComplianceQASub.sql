SELECT
    ComplianceCurrentDateRangeMonths.Year,
    ComplianceCurrentDateRangeMonths.Month,
Val(Replace([MonthNo],"-","")) AS MonNum,
    IIf(
IsNull([COmplianceQA2].[Month]),
        "No Data Submission",
        "Successful Submission"
    )AS [Submit Status]

FROM
    ComplianceQA2
    RIGHT JOIN ComplianceCurrentDateRangeMonths
        ON (ComplianceQA2.Month = ComplianceCurrentDateRangeMonths.Month)
        AND
        (ComplianceQA2.Year = ComplianceCurrentDateRangeMonths.Year)

ORDER BY
    ComplianceCurrentDateRangeMonths.Year
    Val(Replace([MonthNo],"-",""));
