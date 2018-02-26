Attribute VB_Name = "M_0_AddCMRRibbonButtons"
Option Explicit

Const cCommandBar = "CaSES"
Const cCommandBar2 = "CMR Tools 2"
Const cCommandBar3 = "CMR Tools 3"

Sub Toolbar_ON()
    
    'This portion of code runs the ribbon macro to build the CaSES tab on the MS Excel ribbon
    Dim Bar As CommandBar
    Dim bar2 As CommandBar
    Dim bar3 As CommandBar
    
    Dim ctl As CommandBarControl
    Dim bln As Boolean
'    Dim iHelpMenu As Integer
    Dim ModelDropDown As CommandBarControl
    Dim ModelReviewDropDown As CommandBarControl
    Dim ModelPropDropDown As CommandBarControl
    Dim EstimateDropDown As CommandBarControl
    Dim ModelFormatDropDrown As CommandBarControl
    Dim Fixmymodeldropdown As CommandBarControl
    
    On Error Resume Next
    'deletes the toolbar when the addin is closed
    Application.CommandBars(cCommandBar).Delete
    Application.CommandBars(cCommandBar2).Delete
    Application.CommandBars(cCommandBar3).Delete
    On Error GoTo 0
    
    bln = Application.ScreenUpdating
    Application.ScreenUpdating = False

'**********************************************************************************************************************************
'This area of code checks for the first CommandBar in the Addin tab. If it does not exist, the commandbar is added.
' if it already exists, then it is deleted and then recreated
    
    On Error Resume Next: Set Bar = Application.CommandBars(cCommandBar): On Error GoTo 0
    If Bar Is Nothing Then Set Bar = Application.CommandBars.Add(Name:=cCommandBar, Position:=msoBarTop)
       
'**********************************************************************************************************************************
 ' This area of code builds the first dropdown menu shown on the Addin Tab
 'Application.Run "addin.xla!Work"
 
    bln = Application.ScreenUpdating
    Application.ScreenUpdating = False
    With Bar.Controls.Add(Type:=msoControlButton)
        .FaceId = 30
        .Caption = "About CaSES"
        .Style = 3
        .OnAction = ThisWorkbook.Name & "!About_CT"
        .TooltipText = "About CaSES Add-in"
    End With
     
     
''''''''''''''''''''''''''''''''''''''
'Model Template
''''''''''''''''''''''''''''''''''''''

     Set ModelDropDown = Bar.Controls.Add(Type:=msoControlPopup)
                  
     'Give the control1 a caption
     ModelDropDown.Caption = "& Model Template"
     
     With ModelDropDown.Controls.Add(Type:=msoControlButton)
        .Caption = "Open New Model"
        .OnAction = ThisWorkbook.Name & "!OpenModel"
     End With
     
     With ModelDropDown.Controls.Add(Type:=msoControlButton)
        .Caption = "Open Uncertainty Template"
        .OnAction = ThisWorkbook.Name & "!OpenUncertainty"
     End With
     
     With ModelDropDown.Controls.Add(Type:=msoControlButton)
        .Caption = "Open JA CSRUH Example"
        .OnAction = ThisWorkbook.Name & "!Open_JACSRUH"
     End With
         
     
''''''''''''''''''''''''''''''''''''''''
'Audit
''''''''''''''''''''''''''''''''''''''''
     
     Set ModelReviewDropDown = _
     Bar.Controls.Add(Type:=msoControlPopup)
                          
     'Give the control1 a caption
     ModelReviewDropDown.Caption = "& Model Review Toolkit"
        
        With ModelReviewDropDown.Controls.Add(Type:=msoControlButton)
            .FaceId = 26
            .Caption = "Model Comment Tracker (MCT)"
            .Style = 3
            .OnAction = ThisWorkbook.Name & "!Show_CommentTracker"
            .TooltipText = "This will help sum your WBS elements"
        End With
        
        With ModelReviewDropDown.Controls.Add(Type:=msoControlButton)
            .FaceId = 15
            .Caption = "Traceback Navigator Tool (TNT)"
            .Style = 3
            .OnAction = ThisWorkbook.Name & "!Formula_Auditing"
            .TooltipText = "This will help sum your WBS elements"
        End With
        
        
        
        '' Add dropdown within first comstom dropdown
        Dim PPTDropDown As CommandBarControl
        Set PPTDropDown = ModelReviewDropDown.Controls.Add(Type:=msoControlPopup)
            PPTDropDown.Caption = "Convert Excel Chart to PowerPoint"
            With PPTDropDown.Controls.Add(Type:=msoControlButton)
                .Caption = "Convert All Charts to PowerPoint"
                .OnAction = ThisWorkbook.Name & "!M_AllChartsToPPT"
                .FaceId = 17
                .TooltipText = "This tool will convert every chart in current workbook to a new PowerPoint presentation"
            End With
            With PPTDropDown.Controls.Add(Type:=msoControlButton)
                .Caption = "Convert ONLY Chart Sheets to PowerPoint"
                .OnAction = ThisWorkbook.Name & "!pptPasteAllChartsheet"
                .FaceId = 17
                .TooltipText = "This tool will convert all Chart Sheets to PowerPoint. It will NOT convert charts contained within worksheets"
            End With
            With PPTDropDown.Controls.Add(Type:=msoControlButton)
                .Caption = "Convert ONLY this sheet charts toPowerPoint"
                .OnAction = ThisWorkbook.Name & "!pptPasteCurrentCharts"
                .FaceId = 17
                .TooltipText = "This tool will convert ONLY the chart(s) on the current Sheet"
            End With

        '' end Dropdown
        
    Set ModelPropDropDown = _
     ModelReviewDropDown.Controls.Add(Type:=msoControlPopup)
        
     'Give the control2 a caption
     ModelPropDropDown.Caption = "&Model Properties"
          
        With ModelPropDropDown.Controls.Add(Type:=msoControlButton)
            .Caption = "Create Table of Contents (TOC)"
            .OnAction = ThisWorkbook.Name & "!TEST_CreateTOC3"
            .FaceId = 209
        End With
        
        With ModelPropDropDown.Controls.Add(Type:=msoControlButton)
            .Caption = "Get All Cell Comments"
            .OnAction = ThisWorkbook.Name & "!M_Retrieve_AllComments"
            .FaceId = 210
        End With
        
        With ModelPropDropDown.Controls.Add(Type:=msoControlButton)
            .Caption = "Get All Formula Names"
            .OnAction = ThisWorkbook.Name & "!M_Paste_NamesList"
            .FaceId = 211
        End With

     Set ModelFormatDropDrown = _
     ModelReviewDropDown.Controls.Add(Type:=msoControlPopup)
        
     'Give the control4 a caption
     ModelFormatDropDrown.Caption = "&Automatic Model Formatter (AMF) Tool"
     ModelFormatDropDrown.TooltipText = "Color cells based on the cells content. Useful for auditing Models"
     
        With ModelFormatDropDrown.Controls.Add(Type:=msoControlButton)
            .Caption = "Show Formatting Guide"
            .OnAction = ThisWorkbook.Name & "!printGuide"
            .FaceId = 209
        End With
        
        With ModelFormatDropDrown.Controls.Add(Type:=msoControlButton)
            .Caption = "Format Entire Workbook"
            .OnAction = ThisWorkbook.Name & "!formatEntireWorkbook"
            .FaceId = 209
        End With
        
        With ModelFormatDropDrown.Controls.Add(Type:=msoControlButton)
            .Caption = "Format Worksheet"
            .OnAction = ThisWorkbook.Name & "!formatWorksheet"
            .FaceId = 209
        End With

        With ModelFormatDropDrown.Controls.Add(Type:=msoControlButton)
            .Caption = "Format Cells With Out Dependents (Worksheet)"
            .OnAction = ThisWorkbook.Name & "!colorCellsWithOutDependents"
            .FaceId = 209
        End With
        
     Set Fixmymodeldropdown = ModelReviewDropDown.Controls.Add(Type:=msoControlPopup)
     Fixmymodeldropdown.Caption = "Fix My Model"
          
        With Fixmymodeldropdown.Controls.Add(Type:=msoControlButton)
            .Caption = "Show Hidden Names"
            .OnAction = ThisWorkbook.Name & "!M_Unhide_AllNames"
            .FaceId = 201
        End With
        
        With Fixmymodeldropdown.Controls.Add(Type:=msoControlButton)
            .Caption = "Purge Named Ranges"
            .OnAction = ThisWorkbook.Name & "!M_Delete_NamedRange"
            .FaceId = 202
        End With
        
        With Fixmymodeldropdown.Controls.Add(Type:=msoControlButton)
            .Caption = "Break All Links"
            .OnAction = ThisWorkbook.Name & "!M_BreakLinks"
            .FaceId = 207
        End With
        
        With Fixmymodeldropdown.Controls.Add(Type:=msoControlButton)
            .Caption = "Delete Active Array"
            .OnAction = ThisWorkbook.Name & "!M_DeleteActiveArray"
            .FaceId = 207
        End With
        
        With Fixmymodeldropdown.Controls.Add(Type:=msoControlButton)
            .Caption = "Remove Unused Styles"
            .OnAction = ThisWorkbook.Name & "!M_Remove_UnusedStyles"
            .FaceId = 207
        End With
        
        With ModelReviewDropDown.Controls.Add(Type:=msoControlButton)
            .FaceId = 195
            .Caption = "GAO Cost Estimating Criteria"
            .Style = 3
            .OnAction = ThisWorkbook.Name & "!GAO_CriteriaList"
            .TooltipText = "This button provides a quick reference guide to the GAO cost estimating criteria and best practices"
        End With
                
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Estimate Dropdown
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     Set EstimateDropDown = _
     Bar.Controls.Add(Type:=msoControlPopup)
     
     
     'Give the control3 a caption
     EstimateDropDown.Caption = "&Estimating Toolkit"
     
        With EstimateDropDown.Controls.Add(Type:=msoControlButton)
            .Caption = "Add Inflation Worksheet"
            .OnAction = ThisWorkbook.Name & "!copyInflation"
            .FaceId = 422
        End With

        
        '' Add dropdown within first comstom dropdown
        Dim TemplateDropDown As CommandBarControl
        Set TemplateDropDown = EstimateDropDown.Controls.Add(Type:=msoControlPopup)
            TemplateDropDown.Caption = "Add Calculation Template"
            
            With TemplateDropDown.Controls.Add(Type:=msoControlButton)
                .Caption = "Generic Template"
                .OnAction = ThisWorkbook.Name & "!addGenericCalc"
                .FaceId = 215
            End With
            
        '' end Dropdown
        
        Dim WBSDropDown As CommandBarControl
        Set WBSDropDown = EstimateDropDown.Controls.Add(Type:=msoControlPopup)
        WBSDropDown.Caption = "WBS Tool"
            With WBSDropDown.Controls.Add(Type:=msoControlButton)
                .Caption = "Outline WBS Elements"
                .OnAction = ThisWorkbook.Name & "!wbsGroupInd"
                .FaceId = 202
            End With
            
            With WBSDropDown.Controls.Add(Type:=msoControlButton)
                .Caption = "Sum WBS Elements"
                .OnAction = ThisWorkbook.Name & "!sumWBS"
                .FaceId = 201
                .TooltipText = "The WBS MUST use indents in order to work properly. Run WBS Outline  Tool first if your WBS uses periods and NOT indent"
            End With
            
            With WBSDropDown.Controls.Add(Type:=msoControlButton)
                .Caption = "Add WBS to Worksheet"
                .OnAction = ThisWorkbook.Name & "!WBS_MILSTD881C"
                .FaceId = 169
                .TooltipText = "This module will copy a specified WBS to worksheet"
            End With
            
            
            With WBSDropDown.Controls.Add(Type:=msoControlButton)
                .Caption = "Create WBS Tabs"
                .OnAction = ThisWorkbook.Name & "!M_WBSElements_To_Tabs"
                .FaceId = 169
                .TooltipText = "This module will add model template tabs for all selected WBS elements"
            End With
        
        With EstimateDropDown.Controls.Add(Type:=msoControlButton)
            .Caption = "Sum Pivot Fields"
            .TooltipText = "Cycles through all pivot data fields and sets to sum"
            .OnAction = ThisWorkbook.Name & "!PivotFieldsToSum"
            .FaceId = 95
        End With
        
        With EstimateDropDown.Controls.Add(Type:=msoControlButton)
            .Caption = "Flat File Creator"
            .TooltipText = "Automatically creates a flat file output for selected tabs and data content"
            .OnAction = ThisWorkbook.Name & "!Flat_File_Creator"
            .FaceId = 142
        End With 

'**********************************************************************************************************************************
' This portion of code establishes the third CommandBar and support CommandButtons
'    On Error Resume Next: Set bar3 = Application.CommandBars(cCommandBar3): On Error GoTo 0
'    If bar3 Is Nothing Then Set bar3 = Application.CommandBars.Add(Name:=cCommandBar3, Position:=msoBarTop)
       

    
'**********************************************************************************************************************************
'This portion of the code sets the CommandBars to be visible after they are created
    Bar.Visible = True
    'bar2.Visible = True
    'bar3.Visible = True
        
'**********************************************************************************************************************************
'This portion of the code clears any variable that has been previously set to a specified value
    Set Bar = Nothing
    Set bar2 = Nothing
    Set bar3 = Nothing
        
    Set ModelReviewDropDown = Nothing
    Set ModelPropDropDown = Nothing
    Set EstimateDropDown = Nothing
    Set ModelFormatDropDrown = Nothing
    
End Sub

'AdvancedFileProperties = table of contents
'NewList = GetallComments
'DiagramFlipHorizontal = GetFormulaNames
'NotebookLinkCreate = UnhideFormulaNames
'SavePresentationTask = Charts to PPT
'PivotTableClearMenu = Break All Links

'**********************************************************************************************************************************

''Delete Named Range Tools
'    bln = Application.ScreenUpdating
'    Application.ScreenUpdating = False
''    For Each ctl In bar.Controls: ctl.Delete: Next
'    With bar.Controls.Add(Type:=msoControlButton)
'        .FaceId = 2
'        .Caption = "Model Comment Tracker"
'        .Style = 3
'        .OnAction = ThisWorkbook.Name & "!ShowCommentTracker"
'    End With
'
''Retrieve all model comments tool
'    bln = Application.ScreenUpdating
'    Application.ScreenUpdating = False
''    For Each ctl In bar.Controls: ctl.Delete: Next
'    With bar.Controls.Add(Type:=msoControlButton)
'        .FaceId = 204
'        .Caption = "Model Comment Tracker"
'        .Style = 3
'        .OnAction = ThisWorkbook.Name & "!ShowCommentTracker"
'    End With


