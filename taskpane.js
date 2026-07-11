Office.onReady(function () {
  document.getElementById("sendTriggerButton").addEventListener("click", sendTriggerToPowerAutomate);
});

async function sendTriggerToPowerAutomate() {
  const status = document.getElementById("status");
  const button = document.getElementById("sendTriggerButton");
  const projectNumber = document.getElementById("projectNumber").value.trim();

  if (!projectNumber) {
    status.textContent = "חובה להזין מספר פרויקט.";
    return;
  }

  button.disabled = true;
  status.textContent = "שולח טריגר ל-Power Automate...";

  try {
    const item = Office.context.mailbox.item;

    if (!item) {
      throw new Error("לא נמצא מייל פעיל. פתח מייל ואז הפעל את הכפתור.");
    }

    const itemIdEws = item.itemId || "";
    const itemIdRest = getRestItemIdSafe(itemIdEws);

    const payload = {
      triggerName: "SaveContactsFromOutlookAddin",
      projectNumber: projectNumber,

      mail: {
        itemIdEws: itemIdEws,
        itemIdRest: itemIdRest,
        internetMessageId: item.internetMessageId || "",
        subject: item.subject || "",
        conversationId: item.conversationId || "",
        itemType: item.itemType || ""
      },

      outlookUser: {
        email: Office.context.mailbox.userProfile.emailAddress || "",
        displayName: Office.context.mailbox.userProfile.displayName || ""
      },

      source: {
        app: "OutlookAddin",
        button: "Save Contacts",
        sentAtUtc: new Date().toISOString()
      }
    };

    await postToPowerAutomate(payload);

    status.textContent =
      "נשלח ל-Power Automate.\n" +
      "Project: " + projectNumber + "\n" +
      "Subject: " + (payload.mail.subject || "");
  } catch (error) {
    console.error(error);
    status.textContent = "שגיאה: " + error.message;
  } finally {
    button.disabled = false;
  }
}

function getRestItemIdSafe(itemIdEws) {
  if (!itemIdEws) {
    return "";
  }

  try {
    if (Office.context.mailbox.diagnostics.hostName === "OutlookIOS") {
      return itemIdEws;
    }

    return Office.context.mailbox.convertToRestId(
      itemIdEws,
      Office.MailboxEnums.RestVersion.v2_0
    );
  } catch (error) {
    return "";
  }
}

async function postToPowerAutomate(payload) {
  https://defaultac53278e996549ee9675e63f500525.2f.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflo = "PASTE_POWER_AUTOMATE_HTTP_TRIGGER_URL_HERE";

  /*
    נשלח כ-text/plain עם no-cors כדי להימנע מחסימת CORS בדפדפן.
    לכן ב-Power Automate אנחנו עושים:
    json(triggerBody())
  */

  await fetch(flowUrl, {
    method: "POST",
    mode: "no-cors",
    body: JSON.stringify(payload)
  });
}