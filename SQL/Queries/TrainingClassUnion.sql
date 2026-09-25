SELECT
    SDTCOURSECOMP.[Content Title],
    SDTCOURSECOMP.COURSEID,
    SDTCOURSECOMP.[Content Type]

FROM
    SDTCOURSECOMP

GROUP BY
    SDTCOURSECOMP.[Content Title]
    SDTCOURSECOMP.COURSEID
    SDTCOURSECOMP.[Content Type]

HAVING
    (((SDTCOURSECOMP.[Content Type])<>"Curriculum"));
