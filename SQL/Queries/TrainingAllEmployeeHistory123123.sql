SELECT
    [EMP ID] as EMPID,
    "EmployeeOnlyCurrent" as FromTable,
    [POS-START-DATE],
    [Full Name],
    [SEPARATION],
    [ROLE-CODE],
    [JOBTITLE],
    [CC-BASE],
    [SUPV#],
    [ESH-HIRE-DATE],
    [BLDG#],
    [POD],
    [WARD],
    [POSTYP DESC],
    [ESH_subcat],
    [NursingDesignationCode],
    [POSITION#],
    [EMPSTATUS],
    [Supervisor Name]

FROM
    EmployeeOnlyCurrent

UNION ALL
SELECT
    [EMP ID],
    "EMPSERVICE",
    [POS-START-DATE],
    [Full Name],
    [EXIT-DATE],
    [ROLE],
    [JOB TITLE],
    [CC],
    [SUPV#],
    [ESH-HIRE-DATE],
    [BDG],
    [PD],
    [WARD],
    [POSTYP DESC],
    [SUBCAT],
    [DESC],
    [POSITION#],
    [EMPSTATUS],
    [Supervisor Name]

FROM
    TrainingEmpService;
