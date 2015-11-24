Public Interface IFlowSentNotification
    Sub Notify(ByVal dbManager As com.Azion.NET.VB.DatabaseManager, ByVal infoNextStep() As StepInfoItemExt)
End Interface
