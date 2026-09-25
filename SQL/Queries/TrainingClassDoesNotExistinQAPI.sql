SELECT
    SDTCOURSE.COURSEID,
    TrainingClassUnion.[Content Title],
    TrainingClassUnion.COURSEID

FROM
    TrainingClassUnion
    LEFT JOIN SDTCOURSE
        ON TrainingClassUnion.COURSEID = SDTCOURSE.COURSEID

WHERE
    (((SDTCOURSE.COURSEID) Is Null) AND ((TrainingClassUnion.COURSEID) Is Null));
