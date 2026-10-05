/* global Office */
"use strict";

function appointmentParameters(subject, body, now = new Date()) {
  // Reject oversized content instead of silently dropping part of an email.
  if (subject.length > 255) throw new Error("The subject exceeds 255 characters. Shorten it before trying again.");
  if (body.length > 32000) throw new Error("The email body exceeds this add-in's 32,000-character safety limit. Copy the needed portion manually.");
  return {
    subject,
    body,
    start: now,
    end: new Date(now.getTime() + 30 * 60 * 1000)
    // No attendees or resources: Outlook opens an appointment with Save.
  };
}

function createAppointment(event) {
  let item;
  let finished = false;
  function finish(error) {
    if (finished) return;
    finished = true;
    try {
      if (error && item && item.notificationMessages) {
        item.notificationMessages.replaceAsync("appointmentError", {
          type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
          message: String(error.message || error).slice(0, 150)
        }, () => {});
      }
    } catch (_) {
      // A notification failure must not prevent the command from completing.
    } finally {
      event.completed();
    }
  }
  try {
    item = Office.context.mailbox.item;
    if (!item || item.itemType !== Office.MailboxEnums.ItemType.Message) {
      throw new Error("Open an email before using Create Appointment.");
    }
    if (!Office.context.requirements.isSetSupported("Mailbox", "1.9")) {
      throw new Error("This Outlook client needs Mailbox API 1.9 or later.");
    }
    item.body.getAsync(Office.CoercionType.Text, result => {
      try {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          finish(result.error || new Error("Outlook could not read the email body."));
          return;
        }
        const parameters = appointmentParameters(item.subject || "", result.value || "");
        Office.context.mailbox.displayNewAppointmentFormAsync(parameters, opened => {
          finish(opened.status === Office.AsyncResultStatus.Succeeded ? null :
            (opened.error || new Error("Outlook could not open the appointment.")));
        });
      } catch (error) { finish(error); }
    });
  } catch (error) { finish(error); }
}

if (typeof Office !== "undefined") {
  Office.onReady(() => Office.actions.associate("createAppointment", createAppointment));
}
if (typeof module !== "undefined") module.exports = { appointmentParameters, createAppointment };
