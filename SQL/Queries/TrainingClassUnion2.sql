SELECT
    SDTCOURSE.COURSEID,
    SDTCOURSE.COURSENAME

FROM
    TrainingClassUnion
    INNER JOIN SDTCOURSE
        ON TrainingClassUnion.[Content Title] = SDTCOURSE.COURSENAME;
