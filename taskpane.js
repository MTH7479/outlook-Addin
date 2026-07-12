Office.onReady(function () {
  document
    .getElementById("sendTriggerButton")
    .addEventListener("click", sendTriggerToPowerAutomate);
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
  status.textContent = "שולח אנשי קשר ל-Power Automate...";

  try {
    const item = Office.context.mailbox.item;

    if (!item) {
      throw new Error("לא נמצא מייל פעיל. פתח מייל ואז הפעל את הכפתור.");
    }

    const itemIdEws = item.itemId || "";
    const itemIdRest = getRestItemIdSafe(itemIdEws);

    const contacts = collectContactsFromMailItem(item);

    if (contacts.length === 0) {
      throw new Error("לא נמצאו אנשי קשר במייל.");
    }

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

      contacts: contacts,

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

    const responseText = await postToPowerAutomate(payload);

    status.textContent =
      "נשלח ל-Power Automate בהצלחה.\n" +
      "Project: " + projectNumber + "\n" +
      "Contacts: " + contacts.length + "\n" +
      "Subject: " + (payload.mail.subject || "") + "\n" +
      "Response: " + responseText;

  } catch (error) {
    console.error(error);
    status.textContent = "שגיאה: " + error.message;
  } finally {
    button.disabled = false;
  }
}

function collectContactsFromMailItem(item) {
  const contacts = [];

  // From
  if (item.from) {
    contacts.push({
      name: item.from.displayName || "",
      email: normalizeEmail(item.from.emailAddress || ""),
      source: "From"
    });
  }

  // To
  if (Array.isArray(item.to)) {
    item.to.forEach(function (recipient) {
      contacts.push({
        name: recipient.displayName || "",
        email: normalizeEmail(recipient.emailAddress || ""),
        source: "To"
      });
    });
  }

  // Cc
  if (Array.isArray(item.cc)) {
    item.cc.forEach(function (recipient) {
      contacts.push({
        name: recipient.displayName || "",
        email: normalizeEmail(recipient.emailAddress || ""),
        source: "Cc"
      });
    });
  }

  return removeDuplicateContacts(contacts);
}

function normalizeEmail(email) {
  return String(email || "").trim().toLowerCase();
}

function removeDuplicateContacts(contacts) {
  const seen = {};
  const result = [];

  contacts.forEach(function (contact) {
    if (!contact.email) {
      return;
    }

    const key = contact.email;

    if (!seen[key]) {
      seen[key] = true;
      result.push(contact);
    }
  });

  return result;
}

function getRestItemIdSafe(itemIdEws) {
  if (!itemIdEws) {
    return "";
  }

  try {
    return Office.context.mailbox.convertToRestId(
      itemIdEws,
      Office.MailboxEnums.RestVersion.v2_0
    );
  } catch (error) {
    return "";
  }
}

async function postToPowerAutomate(payload) {
  const flowUrl = "https://defaultac53278e996549ee9675e63f500525.2f.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/12/workflows/3b502b33e84c4123a19bc4770cbd31c0/triggers/manual/paths/invoke?api-version=1";

  const response = await fetch(flowUrl, {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify(payload)
  });

  const responseText = await response.text();

  if (!response.ok) {
    throw new Error(
      "Power Automate החזיר שגיאה. Status: " +
      response.status +
      ". Response: " +
      responseText
    );
  }

  return responseText || "OK";
}
