import { saveAs } from 'file-saver';
import { OpenAIClient, AzureKeyCredential } from '@azure/openai';

Office.onReady((info) => {
  // Check if we're in Outlook
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("save-button").onclick = saveEmailAsJson;
    document.getElementById("openai-button").onclick = sendToAzureOpenAI;
  }
});

/**
 * Gets the current email item and extracts its content to save as JSON
 */
function saveEmailAsJson() {
  const statusElement = document.getElementById("status");
  statusElement.innerText = "Processing...";

  try {
    // Get the current item (email)
    const item = Office.context.mailbox.item;
    
    if (!item) {
      statusElement.innerText = "No email selected";
      return;
    }

    // Create an object to store email data
    const emailData = {
      subject: item.subject,
      sender: item.sender ? item.sender.emailAddress : "Unknown",
      receivedTime: item.dateTimeCreated ? new Date(item.dateTimeCreated).toISOString() : null,
      bodyContent: null
    };

    // Get the email body content
    item.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        // Add the body text to our data object
        emailData.bodyContent = result.value;
        
        // Convert the email data to a JSON string
        const jsonData = JSON.stringify(emailData, null, 2);
        
        // Create a Blob from the JSON string
        const blob = new Blob([jsonData], { type: "application/json" });
        
        // Generate a filename using the subject (or a default name if no subject)
        const filename = `${emailData.subject || "email"}_${new Date().getTime()}.json`;
        
        // Save the file
        saveAs(blob, filename);
        
        statusElement.innerText = "Email saved as JSON successfully!";
      } else {
        statusElement.innerText = `Error getting email body: ${result.error.message}`;
      }
    });
  } catch (error) {
    statusElement.innerText = `Error: ${error.message}`;
  }
}

/**
 * Sends email content to Azure OpenAI for processing
 */
async function sendToAzureOpenAI() {
  const statusElement = document.getElementById("status");
  const responseContainer = document.getElementById("response-container");
  
  statusElement.innerText = "Sending to Azure OpenAI...";
  responseContainer.style.display = "none";
  
  try {
    // Get the current item (email)
    const item = Office.context.mailbox.item;
    
    if (!item) {
      statusElement.innerText = "No email selected";
      return;
    }
    
    // Get email data
    item.body.getAsync(Office.CoercionType.Text, async (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const emailData = {
          subject: item.subject,
          sender: item.sender ? item.sender.emailAddress : "Unknown",
          receivedTime: item.dateTimeCreated ? new Date(item.dateTimeCreated).toISOString() : null,
          bodyContent: result.value
        };
        
        // Create prompt for Azure OpenAI
        const prompt = `
        Please analyze the following email and provide a brief summary:
        
        Subject: ${emailData.subject}
        From: ${emailData.sender}
        Date: ${emailData.receivedTime}
        
        Body:
        ${emailData.bodyContent}
        `;
        
        try {
          // Initialize Azure OpenAI client
          const endpoint = process.env.AZURE_OPENAI_ENDPOINT;
          const apiKey = process.env.AZURE_OPENAI_API_KEY;
          const deploymentName = process.env.AZURE_OPENAI_DEPLOYMENT_NAME;
          
          if (!endpoint || !apiKey || !deploymentName) {
            statusElement.innerText = "Azure OpenAI credentials not configured. Please check your .env file.";
            return;
          }
          
          const client = new OpenAIClient(
            endpoint,
            new AzureKeyCredential(apiKey)
          );
          
          // Get response from Azure OpenAI
          const response = await client.getCompletions(
            deploymentName,
            [prompt],
            {
              maxTokens: 800,
              temperature: 0.7,
              topP: 0.95,
              frequencyPenalty: 0,
              presencePenalty: 0,
              stopSequences: ["---"]
            }
          );
          
          // Display the response
          if (response.choices && response.choices.length > 0) {
            const aiResponse = response.choices[0].text.trim();
            responseContainer.innerText = aiResponse;
            responseContainer.style.display = "block";
            statusElement.innerText = "Response received from Azure OpenAI!";
          } else {
            statusElement.innerText = "No response received from Azure OpenAI.";
          }
        } catch (error) {
          console.error("Azure OpenAI API error:", error);
          statusElement.innerText = `Azure OpenAI Error: ${error.message}`;
        }
      } else {
        statusElement.innerText = `Error getting email body: ${result.error.message}`;
      }
    });
  } catch (error) {
    statusElement.innerText = `Error: ${error.message}`;
  }
}

// Fallback function if Office.js is not available
function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal(); 