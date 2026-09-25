SELECT
    TrainingALLEmployeeHistory.*,
    SDTGLAIT.GLAIT

FROM
    SDTGLAIT
    RIGHT JOIN TrainingALLEmployeeHistory
        ON SDTGLAIT.EMPID = TrainingALLEmployeeHistory.EMPID;
