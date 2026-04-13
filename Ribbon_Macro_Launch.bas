Attribute VB_Name = "Ribbon_Macro_Launch"
Sub RPubBoM(control As IRibbonControl)
    Call toolcheck
    Call templateCheck
    Call PubBoM
End Sub
Sub RpubEST(control As IRibbonControl)
    Call toolcheck
    Call templateCheck
    Call pubEST
End Sub
Sub RsumSheetSet(control As IRibbonControl)
    Call toolcheck
    Call templateCheck
    Call excSetup
    Call sumSheetSet
End Sub
Sub RimportBids(control As IRibbonControl)
    Call toolcheck
    Call templateCheck
    Call importBids
End Sub
Sub RnewSys(control As IRibbonControl)
    Call toolcheck
    Call templateCheck
    Call newSys
End Sub
Sub rpCounts(control As IRibbonControl)
    Call toolcheck
    Call templateCheck
    Call pCounts
End Sub
Sub RrevUp(control As IRibbonControl)
    Call toolcheck
    Call templateCheck
    Call revUp
End Sub
Sub RbbRev(control As IRibbonControl)
    Call toolcheck
    Call templateCheck
    Call bbRev
End Sub
Sub RdwgPull(control As IRibbonControl)
    Call toolcheck
    Call templateCheck
    Call dwgPull
End Sub
Sub RSystemRow(control As IRibbonControl)
    Call toolcheck
    Call templateCheck
    Call SystemRow
End Sub
Sub Rautoup(control As IRibbonControl)
'    Call toolcheck
'    Call templateCheck
'    Call autoup
End Sub
Sub RpopulateMASTER(control As IRibbonControl)
'    Call toolcheck
'    Call templateCheck
'    Call populateMASTER
End Sub
Sub RdeleteSys(control As IRibbonControl)
    Call toolcheck
    Call templateCheck
    Call deleteSys
End Sub
Sub RsendtoPDB(control As IRibbonControl)
'    Call toolcheck
'    Call templateCheck
'    Call sendtoPDB
End Sub
Sub RofeRow(control As IRibbonControl)
    Call toolcheck
    Call templateCheck
    Call ofciRow
End Sub
Sub RmArchive(control As IRibbonControl)
    Call toolcheck
    Call templateCheck
    Call mArchive
End Sub
Sub RvCheck(control As IRibbonControl)
    Call toolcheck
    Call vCheck
End Sub
Sub Rpresentationmode(control As IRibbonControl)
    Call toolcheck
    Call templateCheck
    Call presentationmode
End Sub
Sub Rworkmode(control As IRibbonControl)
    Call toolcheck
    Call templateCheck
    Call workmode
End Sub
Sub RtrackC()
    Call toolcheck
    Call templateCheck
    Call trackC
End Sub
Sub RstandRow(control As IRibbonControl)
    Call toolcheck
    Call templateCheck
    Call standRow
End Sub
Sub RnoteRow(control As IRibbonControl)
    Call toolcheck
    Call templateCheck
    Call noteRow
End Sub
