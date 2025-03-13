Sub RunAllPreparationMacros()
    ' Run macros in a specific order
    
    ' Start with deleting unmatched rows
    Call DeleteUnmatchedRows
    
    ' Clear FPM Display Only
    Call ClearFPMDisplayOnly
    
    ' Consolidate and highlight Plan 1
    Call ConsolidatedSortAndHighlightPlan1
    
    ' Consolidate and highlight Plan 2
    Call ConsolidatedSortAndHighlightPlan2
    
    ' Highlight overdue and next steps
    Call HighlightOverdueAndNextSteps
    
    ' Create PM tabs and distribute data
    Call CreatePMTabsAndDistributeData
    
    ' Final sort across all sheets
    Call FinalSort
    
    ' Optional: Add a message to confirm completion
    MsgBox "All preparation macros have been completed successfully!", vbInformation
End Sub

