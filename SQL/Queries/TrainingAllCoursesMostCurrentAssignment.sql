SELECT
    TrainingAllCourses.COURSEID,
    TrainingAllCourses.COURSEIDNUM,
    TrainingAllCourses.COURSENAME,
    TrainingAllCourses.TRNCATID2,
    TrainingMostCurrentClassAssignment.MaxOfDATEASSIGN,
    SDTASSIGNCLASS.DATEDUE

FROM
    TrainingAllCourses
    LEFT JOIN (TrainingMostCurrentClassAssignment LEFT JOIN SDTASSIGNCLASS
        ON (TrainingMostCurrentClassAssignment.COURSEID = SDTASSIGNCLASS.COURSEID)
        AND
        (TrainingMostCurrentClassAssignment.[MaxOfDATEASSIGN] = SDTASSIGNCLASS.DATEASSIGN)) ON TrainingAllCourses.COURSEID = TrainingMostCurrentClassAssignment.COURSEID;
