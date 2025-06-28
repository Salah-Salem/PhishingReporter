/* eslint-disable no-undef */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    // document.getElementById("run").onclick = run;
    document.getElementById("run").onclick = reportPhishing;
  }
});

// export async function run() {
//   /**
//    * Insert your Outlook code here
//    */
//   Office.context.mailbox.displayNewMessageForm({
//     toRecipients: [process.env.SUPPORT_EMAIL_ADDRESS],
//     subject: process.env.COMPANY_NAME + " - Phishing Report ( Subject: " + Office.context.mailbox.item.subject + " )",
//     htmlBody: "Hello,<br><br>I believe the attached email is a scam or phishing email.<br><br>Thanks.",
//     attachments: [
//       {
//         type: "item",
//         itemId: Office.context.mailbox.item.itemId,
//         name: Office.context.mailbox.item.subject + ".msg",
//       },
//     ],
//   });
// }

// Helper function to get email body

// function getBodyAsync() {
//   return new Promise((resolve, reject) => {
//     Office.context.mailbox.item.body.getAsync("text", { asyncContext: "This is passed to the callback" }, (result) => {
//       if (result.status === Office.AsyncResultStatus.Succeeded) {
//         resolve(result.value);
//       } else {
//         reject(new Error("Failed to get email body"));
//       }
//     });
//   });
// }

// Get email body as HTML
function getHtmlBodyAsync() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.getAsync(
      Office.CoercionType.Html,
      { asyncContext: "This is passed to the callback" },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject(new Error("Failed to get email HTML body"));
        }
      }
    );
  });
}

// Helper function to get email as EML file
// function getItemAsEml() {
//   return new Promise((resolve) => {
//     if (Office.context.mailbox.item.saveAsync) {
//       Office.context.mailbox.item.saveAsync(Office.MailboxEnums.ItemSaveFormat.Eml, (result) => {
//         if (result.status === Office.AsyncResultStatus.Succeeded) {
//           // Convert the base64 string to a Blob
//           const byteCharacters = atob(result.value);
//           const byteNumbers = new Array(byteCharacters.length);
//           for (let i = 0; i < byteCharacters.length; i++) {
//             byteNumbers[i] = byteCharacters.charCodeAt(i);
//           }
//           const byteArray = new Uint8Array(byteNumbers);
//           const blob = new Blob([byteArray], { type: "message/rfc822" });
//           resolve(blob);
//         } else {
//           console.warn("Failed to save email as EML");
//           resolve(null);
//         }
//       });
//     } else {
//       console.warn("saveAsync not supported in this version of Outlook");
//       resolve(null);
//     }
//   });
// }

async function reportPhishing() {
  try {
    const item = Office.context.mailbox.item;
    const statusElement = document.getElementById("status-message");
    const body = await getHtmlBodyAsync();

    const emailData = {
      date: new Date().toISOString().split("T")[0],
      subject: item.subject || "",
      sender_name: item.sender.displayName || "",
      sender_email: item.sender.emailAddress || "",
      content: body || "",
      authentication: "",
      link: "",
      evaluation: "",
      action_status: "",
    };
    // Show loading state
    document.getElementById("run").disabled = true;
    statusElement.style.display = "block";
    statusElement.textContent = "Collecting email information...";
    statusElement.style.color = "black";

    statusElement.textContent = "Submitting report...";

    // Submit to your API
    const response = await submitToApi(emailData);

    if (response.success) {
      statusElement.textContent = "Report submitted successfully!";
      statusElement.style.color = "green";
    } else {
      statusElement.textContent = "Error: " + (response.message || "Failed to submit report");
      statusElement.style.color = "red";
    }
  } catch (error) {
    console.error("Error reporting phishing:", error);
    const statusElement = document.getElementById("status-message");
    statusElement.style.display = "block";
    statusElement.textContent = "Error: " + (error.message || "Failed to submit report");
    statusElement.style.color = "red";
  } finally {
    document.getElementById("run").disabled = false;
  }
}

async function submitToApi(emailData) {
  // Replace with your actual API endpoint
  const apiUrl = "https://resilience-stag-api.orgate.io/api/v1/store-report-logs";
  try {
    const response = await fetch(apiUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Accept: "application/json",
        // Add any required authentication headers
      },
      body: JSON.stringify(emailData),
    });
    if (!response.ok) {
      throw new Error(`API request failed with status ${response.status}`);
    }

    return await response.json();
  } catch (error) {
    console.error("API request failed:", error);
    throw error;
  }
}
