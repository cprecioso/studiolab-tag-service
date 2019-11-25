const TEMPLATE_ID = "1sY6kwlYECcVAII7bgOoe9dRr9xcy0I9uCKhaWyAHfSM"
const ATTACHMENT_FILENAME = "StudioLab Tag"
const MAIL_PARAMETERS: Partial<GoogleAppsScript.Mail.MailAdvancedParameters> = {
  name: "StudioLab Tag",
  subject: "Your new tag",
  body: `Here's your prototype tag. Please attach it to your prototype.`
}

const submitTriggerHandler = onSubmit.name

function onOpen() {
  FormApp.getUi()
    .createMenu("PDF Tag Creator")
    .addItem("Make PDF tag for response", onManualTrigger.name)
    .addSeparator()
    .addItem("Start automatic tag making", onConfigure.name)
    .addItem("Stop automatic tag making", onUnconfigure.name)
    .addToUi()
}

function onConfigure() {
  ScriptApp.newTrigger(submitTriggerHandler)
    .forForm(FormApp.getActiveForm())
    .onFormSubmit()
    .create()
}

function onUnconfigure() {
  const trigger = getTrigger()
  if (!trigger) throw new Error("No trigger found")
  ScriptApp.deleteTrigger(trigger)
}

function getTrigger() {
  for (const trigger of ScriptApp.getUserTriggers(FormApp.getActiveForm())) {
    if (
      trigger.getEventType() === ScriptApp.EventType.ON_FORM_SUBMIT &&
      trigger.getHandlerFunction() === submitTriggerHandler
    )
      return trigger
  }
  return false
}

function onManualTrigger() {
  const form = FormApp.getActiveForm()
  const ui = FormApp.getUi()

  const promptResponse = ui.prompt(
    "Enter the Response number",
    ui.ButtonSet.OK_CANCEL
  )
  if (promptResponse.getSelectedButton() === ui.Button.CANCEL) return

  const responseNumber = parseInt(promptResponse.getResponseText(), 10)
  if (isNaN(responseNumber)) {
    ui.alert("That's not a number")
    return
  }

  const response = form.getResponses()[responseNumber - 1]
  if (!response) {
    ui.alert("Can't find the given ID")
    return
  }

  const data = extractData(response)
  const pdf = makePDF(data)

  const pdfFile = DriveApp.createFile(pdf)
  pdfFile.setTrashed(true)
  const pdfUrl = pdfFile.getUrl()

  const htmlOutput = HtmlService.createHtmlOutput(
    `<a target="_blank" href="${pdfUrl}">View PDF</a>`
  )
  ui.showModalDialog(htmlOutput, "Done")
}

function onSubmit(event: GoogleAppsScript.Events.FormsOnFormSubmit) {
  const email = event.response.getRespondentEmail()
  if (!email) throw new Error("No email found")

  const data = extractData(event.response)
  const pdf = makePDF(data)

  MailApp.sendEmail({
    ...MAIL_PARAMETERS,
    to: email,
    body:
      MAIL_PARAMETERS.body +
      `\n\nUse this link to modify your prototype information: ${event.response.getEditResponseUrl()}`,
    attachments: [pdf]
  })
}

function extractData(
  response: GoogleAppsScript.Forms.FormResponse
): { name: string; value: string }[] {
  const data = response.getItemResponses().map(itemResponse => ({
    name: "" + itemResponse.getItem().getTitle(),
    value: "" + itemResponse.getResponse()
  }))

  const email = response.getRespondentEmail()
  if (email) data.push({ name: "Email", value: email })

  const editURL = response.getEditResponseUrl()
  if (editURL) data.push({ name: "Edit URL", value: editURL })

  return data
}

function makePDF(data: { name: string; value: string }[]) {
  const template = DriveApp.getFileById(TEMPLATE_ID)
  const copyFile = template.makeCopy(ATTACHMENT_FILENAME)
  const copyId = copyFile.getId()

  const copySlide = SlidesApp.openById(copyId)
  for (const { name, value } of data) {
    copySlide.replaceAllText(`%${name}%`, value)
  }

  copySlide.saveAndClose()
  copyFile.setTrashed(true)

  const blob = copyFile.getAs("application/pdf")
  blob.setName(ATTACHMENT_FILENAME + ".pdf")
  return blob
}
