function getTriggerInfo() {

    const triggerInfo = ScriptApp.getProjectTriggers().map(trigger => [trigger.getUniqueId(), trigger.getEventType(), trigger.getHandlerFunction()])
    Logger.log(triggerInfo) 
}

function deleteaddUuidAndEmailCheckboxTrigger() {
    const triggers = ScriptApp.getProjectTriggers()

    const triggerToDelete = triggers.find(trigger => trigger.getUniqueId() == '6547125025935885635') 

    Logger.log(triggerToDelete.getHandlerFunction())

    ScriptApp.deleteTrigger(triggerToDelete)
}