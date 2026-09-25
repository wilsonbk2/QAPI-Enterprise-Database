SELECT
    TrainingRequiredClassesbyRole.*,
    TrainingStaffWithGLAITs.[EMP ID],
    TrainingStaffWithGLAITs.GLAIT,
    TrainingStaffWithGLAITs.[SUPV#],
    TrainingStaffWithGLAITs.SEPARATION,
    TrainingStaffWithGLAITs.EMPSTATUS,
    TrainingStaffWithGLAITs.[ESH-HIRE-DATE],
    TrainingStaffWithGLAITs.[POS-START-DATE]

FROM
    TrainingRequiredClassesbyRole
    INNER JOIN TrainingStaffWithGLAITs
        ON TrainingRequiredClassesbyRole.[ROLE-CODE] = TrainingStaffWithGLAITs.[ROLE-CODE];
