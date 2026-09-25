SELECT
    UtilizationAssessmentIntRecords.PatientID,
    UtilizationAssessmentIntRecords.EPI,
    UtilizationAssessmentIntRecords.[Patient Name],
    UtilizationAssessmentIntRecords.BUI,
    UtilizationAssessmentIntRecords.UNIT,
    [ADMISDT]+[ADMISTM] AS [Admission D/T],
    UtilizationAssessmentIntRecords.[Psychosocial Code],
    DateAdd(
        "h",
        [Psychosocial Hours],
        [Admission D/T]
    )AS PsychosocialNeedDT,
    UtilizationAssessmentCompletion.DateTime AS PsychosocialCompDT,
Val("168") AS [Psychosocial Hours],
Now() AS TodaysDT,
    IIf(
IsNull([PsychosocialCompDT]),
        IIf(
            [TodaysDT]>[PsychosocialNeedDT],
            "Incomplete",
            "Mark Complete"
        ),
        IIf(
            [PsychosocialCompDT]<=[PsychosocialNeedDT],
            "Complete",
            "Late"
        )
    )AS Psychosocial,
    UtilizationAssessmentIntRecords.[Px Rec Code],
    DateAdd(
        "h",
        [PxRecHours],
        [Admission D/T]
    )AS PxRecNeedDT,
Val("168") AS PxRecHours,
    UtilizationAssessmentCompletion_2.DateTime AS PxRecCompDT,
    IIf(
IsNull([PxRecCompDT]),
        IIf(
            [TodaysDT]>[PxRecNeedDT],
            "Incomplete",
            "Mark Complete"
        ),
        IIf(
            [PxRecCompDT]<=[PxRecNeedDT],
            "Complete",
            "Late"
        )
    )AS [Px Rec],
    UtilizationAssessmentIntRecords.[Rehab CorRehab Code],
    DateAdd(
        "h",
        [RehabHours],
        [Admission D/T]
    )AS RehabNeedDT,
Val("168") AS RehabHours,
    UtilizationAssessmentCompletion_1.DateTime AS RehabCompDT,
    IIf(
IsNull([RehabCompDT]),
        IIf(
            [TodaysDT]>[RehabNeedDT],
            "Incomplete",
            "Mark Complete"
        ),
        IIf(
            [RehabCompDT]<=[RehabNeedDT],
            "Complete",
            "Late"
        )
    )AS Rehab,
    UtilizationAssessmentIntRecords.Nutrition AS [Nutrition Code],
    DateAdd(
        "h",
        [NutritionHours],
        [Admission D/T]
    )AS NutritionNeedDT,
Val("168") AS NutritionHours,
    UtilizationAssessmentCompletion_3.DateTime AS NutritionCompDT,
    IIf(
IsNull([NutritionCompDT]),
        IIf(
            [TodaysDT]>[NutritionNeedDT],
            "Incomplete",
            "Mark Complete"
        ),
        IIf(
            [NutritionCompDT]<=[NutritionNeedDT],
            "Complete",
            "Late"
        )
    )AS Nutrition,
    UtilizationAssessmentIntRecords.Dysphagia AS [Dysphagia Code],
    DateAdd(
        "h",
        [DysphagiaHours],
        [Admission D/T]
    )AS DysphagiaNeedDT,
Val("72") AS DysphagiaHours,
    UtilizationAssessmentCompletion_4.DateTime AS DysphagiaCompDT,
    IIf(
IsNull([DysphagiaCompDT]),
        IIf(
            [TodaysDT]>[DysphagiaNeedDT],
            "Incomplete",
            "Mark Complete"
        ),
        IIf(
            [DysphagiaCompDT]<=[DysphagiaNeedDT],
            "Complete",
            "Late"
        )
    )AS Dysphagia,
    UtilizationAssessmentIntRecords.[Falls Code],
    DateAdd(
        "h",
        [FallsHours],
        [Admission D/T]
    )AS FallsNeedDT,
Val("168") AS FallsHours,
    UtilizationAssessmentCompletion_5.DateTime AS FallsCompDT,
    IIf(
IsNull([FallsCompDT]),
        IIf(
            [TodaysDT]>[FallsNeedDT],
            "Incomplete",
            "Mark Complete"
        ),
        IIf(
            [FallsCompDT]<=[FallsNeedDT],
            "Complete",
            "Late"
        )
    )AS Falls,
    UtilizationAssessmentIntRecords.[H/P Code],
    DateAdd(
        "h",
        [H/PHours],
        [Admission D/T]
    )AS [H/PNeedDT],
Val("168") AS [H/PHours],
    UtilizationAssessmentCompletion_6.DateTime AS [H/PCompDT],
    IIf(
IsNull([H/PCompDT]),
        IIf(
            [TodaysDT]>[H/PNeedDT],
            "Incomplete",
            "Mark Complete"
        ),
        IIf(
            [H/PCompDT]<=[H/PNeedDT],
            "Complete",
            "Late"
        )
    )AS [H/P],
    UtilizationAssessmentIntRecords.[Psychiatric Code],
    DateAdd(
        "h",
        [PsychiatricHours],
        [Admission D/T]
    )AS PsychiatricNeedDT,
Val("60") AS PsychiatricHours,
    UtilizationAssessmentCompletion_7.DateTime AS PsychiatricCompDT,
    IIf(
IsNull([PsychiatricCompDT]),
        IIf(
            [TodaysDT]>[PsychiatricNeedDT],
            "Incomplete",
            "Mark Complete"
        ),
        IIf(
            [PsychiatricCompDT]<=[PsychiatricNeedDT],
            "Complete",
            "Late"
        )
    )AS Psychiatric,
    UtilizationAssessmentIntRecords.[Psychology Code],
    DateAdd(
        "h",
        [PsychologyHours],
        [Admission D/T]
    )AS PsychologyNeedDT,
Val("168") AS PsychologyHours,
    UtilizationAssessmentCompletion_8.DateTime AS PsychologyCompDT,
    IIf(
IsNull([PsychologyCompDT]),
        IIf(
            [TodaysDT]>[PsychologyNeedDT],
            "Incomplete",
            "Mark Complete"
        ),
        IIf(
            [PsychologyCompDT]<=[PsychologyNeedDT],
            "Complete",
            "Late"
        )
    )AS Psychology,
    UtilizationAssessmentIntRecords.[Nursing Triage Code],
    DateAdd(
        "h",
        [TriageHours],
        [Admission D/T]
    )AS TriageNeedDT,
Val(0.3) AS TriageHours,
    UtilizationAssessmentCompletion_9.DateTime AS TriageCompDT,
    IIf(
IsNull([TriageCompDT]),
        IIf(
            [TodaysDT]>[TriageNeedDT],
            "Incomplete",
            "Mark Complete"
        ),
        IIf(
            [TriageCompDT]<=[TriageNeedDT],
            "Complete",
            "Late"
        )
    )AS [Nursing Triage],
    UtilizationAssessmentIntRecords.[BH Intake Code],
    DateAdd(
        "h",
        [BHHours],
        [Admission D/T]
    )AS BHNeedDT,
Val("24") AS BHHours,
    UtilizationAssessmentCompletion_10.DateTime AS BHCompDT,
    IIf(
IsNull([BHCompDT]),
        IIf(
            [TodaysDT]>[BHNeedDT],
            "Incomplete",
            "Mark Complete"
        ),
        IIf(
            [BHCompDT]<=[BHNeedDT],
            "Complete",
            "Late"
        )
    )AS [BH Intake],
    Format(
        [ADMISDT],
        "yyyy mmmm"
    )AS MonthYear,
    UtilizationAssessmentIntRecords.patepi

FROM
    ((((((((((UtilizationAssessmentIntRecords LEFT JOIN UtilizationAssessmentCompletion ON (UtilizationAssessmentIntRecords.[Psychosocial Code] = UtilizationAssessmentCompletion.DEQID) AND (UtilizationAssessmentIntRecords.EPI = UtilizationAssessmentCompletion.Episode) AND (UtilizationAssessmentIntRecords.PatientID = UtilizationAssessmentCompletion.PatientID)) LEFT JOIN UtilizationAssessmentCompletion AS UtilizationAssessmentCompletion_2 ON (UtilizationAssessmentIntRecords.[Px Rec Code] = UtilizationAssessmentCompletion_2.DEQID) AND (UtilizationAssessmentIntRecords.EPI = UtilizationAssessmentCompletion_2.Episode) AND (UtilizationAssessmentIntRecords.PatientID = UtilizationAssessmentCompletion_2.PatientID)) LEFT JOIN UtilizationAssessmentCompletion AS UtilizationAssessmentCompletion_1 ON (UtilizationAssessmentIntRecords.[Rehab CorRehab Code] = UtilizationAssessmentCompletion_1.DEQID) AND (UtilizationAssessmentIntRecords.EPI = UtilizationAssessmentCompletion_1.Episode) AND (UtilizationAssessmentIntRecords.PatientID = UtilizationAssessmentCompletion_1.PatientID)) LEFT JOIN UtilizationAssessmentCompletion AS UtilizationAssessmentCompletion_3 ON (UtilizationAssessmentIntRecords.Nutrition = UtilizationAssessmentCompletion_3.DEQID) AND (UtilizationAssessmentIntRecords.EPI = UtilizationAssessmentCompletion_3.Episode) AND (UtilizationAssessmentIntRecords.PatientID = UtilizationAssessmentCompletion_3.PatientID)) LEFT JOIN UtilizationAssessmentCompletion AS UtilizationAssessmentCompletion_4 ON (UtilizationAssessmentIntRecords.Dysphagia = UtilizationAssessmentCompletion_4.DEQID) AND (UtilizationAssessmentIntRecords.EPI = UtilizationAssessmentCompletion_4.Episode) AND (UtilizationAssessmentIntRecords.PatientID = UtilizationAssessmentCompletion_4.PatientID)) LEFT JOIN UtilizationAssessmentCompletion AS UtilizationAssessmentCompletion_5 ON (UtilizationAssessmentIntRecords.[Falls Code] = UtilizationAssessmentCompletion_5.DEQID) AND (UtilizationAssessmentIntRecords.EPI = UtilizationAssessmentCompletion_5.Episode) AND (UtilizationAssessmentIntRecords.PatientID = UtilizationAssessmentCompletion_5.PatientID)) LEFT JOIN UtilizationAssessmentCompletion AS UtilizationAssessmentCompletion_6 ON (UtilizationAssessmentIntRecords.[H/P Code] = UtilizationAssessmentCompletion_6.DEQID) AND (UtilizationAssessmentIntRecords.EPI = UtilizationAssessmentCompletion_6.Episode) AND (UtilizationAssessmentIntRecords.PatientID = UtilizationAssessmentCompletion_6.PatientID)) LEFT JOIN UtilizationAssessmentCompletion AS UtilizationAssessmentCompletion_7 ON (UtilizationAssessmentIntRecords.[Psychiatric Code] = UtilizationAssessmentCompletion_7.DEQID) AND (UtilizationAssessmentIntRecords.EPI = UtilizationAssessmentCompletion_7.Episode) AND (UtilizationAssessmentIntRecords.PatientID = UtilizationAssessmentCompletion_7.PatientID)) LEFT JOIN UtilizationAssessmentCompletion AS UtilizationAssessmentCompletion_8 ON (UtilizationAssessmentIntRecords.[Psychology Code] = UtilizationAssessmentCompletion_8.DEQID) AND (UtilizationAssessmentIntRecords.EPI = UtilizationAssessmentCompletion_8.Episode) AND (UtilizationAssessmentIntRecords.PatientID = UtilizationAssessmentCompletion_8.PatientID)) LEFT JOIN UtilizationAssessmentCompletion AS UtilizationAssessmentCompletion_9 ON (UtilizationAssessmentIntRecords.[Nursing Triage Code] = UtilizationAssessmentCompletion_9.DEQID) AND (UtilizationAssessmentIntRecords.EPI = UtilizationAssessmentCompletion_9.Episode) AND (UtilizationAssessmentIntRecords.PatientID = UtilizationAssessmentCompletion_9.PatientID))
    LEFT JOIN UtilizationAssessmentCompletion AS UtilizationAssessmentCompletion_10
        ON (UtilizationAssessmentIntRecords.[BH Intake Code] = UtilizationAssessmentCompletion_10.DEQID)
        AND
        (UtilizationAssessmentIntRecords.EPI = UtilizationAssessmentCompletion_10.Episode)
        AND
        (UtilizationAssessmentIntRecords.PatientID = UtilizationAssessmentCompletion_10.PatientID);
