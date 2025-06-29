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
  const apiUrl = `${process.env.API_URL}/store-report-logs`;
  console.log("Submitting to API:", apiUrl);
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
