SELECT
    StaffingHoursReport.JOBTITLE,
    StaffingHoursReport.DESC,
    StaffingHoursReport.[Full Name],
    StaffingHoursReport.EMPID,
    StaffingHoursReport.CSCDESC,
    StaffingHoursReport.UNITSUB,
Sum(StaffingHoursReport.THOURS) AS [Total Of THOURS]

FROM
    StaffingHoursReport

GROUP BY
    StaffingHoursReport.JOBTITLE
    StaffingHoursReport.DESC
    StaffingHoursReport.[Full Name]
    StaffingHoursReport.EMPID
    StaffingHoursReport.CSCDESC
    StaffingHoursReport.UNITSUB PIVOT StaffingHoursReport.SHIFT In ("DAY","EVENING","NIGHT");
