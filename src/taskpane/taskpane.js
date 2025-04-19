import { saveAs } from 'file-saver';

Office.onReady((info) => {
  // Check if we're in Outlook
  if (info.host === Office.HostType.Outlook) {
    // Add event handlers for the buttons
    document.getElementById("save-button").onclick = saveEmailAsJson;
    document.getElementById("openai-button").onclick = sendToAzureOpenAI;
    document.getElementById("reply-button").onclick = replyWithResponse;
    
    // Hide the reply button initially - it will be shown when a response is generated
    document.getElementById("reply-button").style.display = "none";
    
    // Log initialization
    console.log("Add-in initialized in Outlook");
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
 * Sends email content to Azure OpenAI for processing using fetch API
 */
async function sendToAzureOpenAI() {
  const statusElement = document.getElementById("status");
  const responseContainer = document.getElementById("response-container");
  const copyButton = document.getElementById("copy-button");
  const replyButton = document.getElementById("reply-button");
  
  statusElement.innerText = "Preparing to send to Azure OpenAI...";
  responseContainer.style.display = "none";
  copyButton.style.display = "none";
  replyButton.style.display = "none";

  try {
    // Get the current item (email)
    const item = Office.context.mailbox.item;
    
    if (!item) {
      statusElement.innerText = "No email selected";
      return;
    }
    
    // Get email data
    item.body.getAsync(Office.CoercionType.Text, async (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        statusElement.innerText = `Error getting email body: ${result.error.message}`;
        return;
      }
      
      try {
        statusElement.innerText = "Processing email content...";
        
        const emailData = {
          subject: item.subject,
          sender: item.sender ? item.sender.emailAddress : "Unknown",
          receivedTime: item.dateTimeCreated ? new Date(item.dateTimeCreated).toISOString() : null,
          bodyContent: result.value
        };
        
        // Hard-coded endpoint for testing
        const endpoint = "https://epmfl.openai.azure.com/openai/deployments/gpt-4o/chat/completions?api-version=2024-02-15-preview";
        
        // Try to get API key from various sources
        let apiKey;
        try {
          // First check window.__env from runtime-config.js
          if (window.__env && window.__env.AZURE_OPENAI_API_KEY) {
            apiKey = window.__env.AZURE_OPENAI_API_KEY;
            console.log("Using API key from runtime config");
          } 
          // Fallback to process.env if available
          else if (typeof process !== 'undefined' && process.env && process.env.AZURE_OPENAI_API_KEY) {
            apiKey = process.env.AZURE_OPENAI_API_KEY;
            console.log("Using API key from process.env");
          } 
          // Final fallback
          else {
            apiKey = "your_api_key_here";
            console.log("No API key found in runtime config or process.env");
          }
        } catch (e) {
          console.error("Error accessing API key:", e);
          apiKey = "your_api_key_here";
        }
        
        if (apiKey === "your_api_key_here") {
          statusElement.innerText = "Azure OpenAI API key not configured. Please add your key to the .env file.";
          return;
        }
        
        console.log("Endpoint URL:", endpoint);
        statusElement.innerText = "Generating legal response...";
        
        // Prepare the request payload for chat completions
        const payload = {
          messages: [
            {
              role: "system",
              content: "You are an AI assistant specialized in analyzing legal correspondence and generating appropriate follow-up emails. Your task is to analyze an email from Bhavesh Mistry (BM) to a client and generate a draft response following specific templates and guidelines. Format your response using markdown for bold text (**bold**) and hyperlinks in the format [link text](URL)."
            },
            {
              role: "user",
              content: `AI Prompt for Email Analysis and Response Generation

Input Format
You will receive: 
1. Initial Email: The original email sent by BM to the client including the subject line
2. Available Meeting Times: A list of dates and times when BM is available for meetings 
3. Templates: Reference email templates to follow

Here is the email to analyze:
Subject: ${emailData.subject}
From: ${emailData.sender}
Date: ${emailData.receivedTime}
Body:
${emailData.bodyContent}

Available Meeting Times:
- Thursday, 27 March 2025 at 10:30am
- Thursday, 27 March 2025 at 1:30pm
- Thursday, 27 March 2025 at 2:30pm

Analysis Requirements
Please analyze the initial email to identify:
1.         Client Information:
–          Extract all client names mentioned in the email greeting (e.g., "Hi Barry" → "Barry")
–          Note if multiple clients are addressed (e.g., couples, business partners)
2.         Service Type:
–          Determine if the email is about estate planning or non-estate planning
–          Estate planning indicators include: "estate planning," "asset protection," "will," "enduring power of attorney," "estate plan," "SMSF Trust Deeds," "Family Trust Deeds"
–          Non-estate planning might relate to: divorce, business matters, disputes, etc.
3.         Meeting Format:
–          Identify if the meeting is proposed as:
•           MS Teams meeting (if call then also MS Teams meeting unless noted specifically that call will be on mobile)
•           In-person meeting at the office
•           In-person meeting at another location (specify if mentioned)
–          Default to "MS Teams meeting OR in-person meeting at our office" if unclear

Response Generation Guidelines
Based on your analysis:
1.         Template Selection:
–          Use Template B if the matter involves estate planning
–          Use Template A for all other matters
2.         Email Structure:
–          Begin with "Private and Confidential"
–          Address the client by name
–          Include "Conference" section with available meeting times
–          For estate planning, include "Questionnaire" section with link
–          End with standard closing and "Reply All" instruction
3.         Formatting Requirements:
–          Maintain the exact formatting from the templates
–          Present meeting times as bullet points
–          Preserve all bold formatting for headings

Template Reference
Template A (Non-Estate Planning)
**Private and Confidential**
 
Hi [client name]
 
Further to Bhavesh Mistry's email correspondence today, we look forward to assisting you.
 
**Conference**
 
Please note, Bhavesh Mistry will be available at the following dates and times below for a **[meeting format]** with you for further discussion:
 
- [date/time option 1]; or
- [date/time option 2]; or
- [date/time option 3].
 
We look forward to hearing from you shortly.
 
Please email us ensuring that you select \`Reply All\` to our email so our team can assist you.

Template B (Estate Planning)
**Private and Confidential**
 
Hi [client name]
 
Further to Bhavesh Mistry's email correspondence today, we look forward to assisting you with your estate planning.
 
**Conference**
 
To allow us to understand your intentions, Bhavesh Mistry will be available at the following dates and times below for a **[meeting format]** with you for an initial discussion of your estate planning:
 
- [date/time option 1]; or
- [date/time option 2]; or
- [date/time option 3].
 
Please kindly let us know if any of the above times are suitable and your best contact number. Alternatively, please let us know if there are any other dates and times more suitable for you.
 
**Questionnaire**
 
In preparation for our conference and to assist us in obtaining your initial information and allowing you to start considering your estate plan, please take a few minutes to complete our estate planning questionnaire at the following [link](https://mistryfallahi.com.au/client-asset-protection-enquiry/).
 
We look forward to hearing from you shortly.
 
Please email us ensuring that you select \`Reply All\` to our email so our team can assist you.

Output Format
Your response should include ONLY the Draft Email which is completely formatted following the appropriate template. Do not include headings like "Analysis" or "Draft Email".

Example output structure:
DRAFT EMAIL:
**Subject: Conference - [Include the complete subject here]**
[Complete formatted email]

Important Notes
•           Do not include section headings like "Analysis:" or "Draft Email:" in your response.
•           Do not include placeholders in the final draft email. All fields should be properly populated.
•           Maintain exact formatting from templates including bold text, bullet points, and paragraph spacing.
•           If anything is unclear, default to the most conservative option and note your uncertainty in the analysis.
•           The email subject line MUST have "Conference -" added at the beginning, followed by the complete original subject.
•           For example: "Conference - Estate Planning - Paulina and Alex Pavlova (Our Ref: 10-0204)" 
•           Date and time to be in this format: Thursday, 27 March 2025 at 10:30am, 1:30pm or 2:30pm.
•           Use **bold** markdown formatting for headings and important text.
•           Format the link in the questionnaire section as a proper markdown link [link](URL).`
            }
          ],
          max_tokens: 1000,
          temperature: 0.2,
          top_p: 0.95,
          frequency_penalty: 0,
          presence_penalty: 0
        };
        
        console.log("Request payload:", JSON.stringify(payload).substring(0, 200) + "...");
        
        // Make the API call
        try {
          const response = await fetch(endpoint, {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
              'api-key': apiKey
            },
            body: JSON.stringify(payload)
          });
          
          console.log("Response status:", response.status, response.statusText);
          
          if (!response.ok) {
            let errorMessage;
            try {
              const errorText = await response.text();
              console.error("Error response:", errorText);
              
              try {
                const errorData = JSON.parse(errorText);
                errorMessage = errorData.error?.message || response.statusText;
              } catch (parseError) {
                errorMessage = `${response.status} ${response.statusText}. ${errorText}`;
              }
            } catch (textError) {
              errorMessage = `${response.status} ${response.statusText}`;
            }
            
            statusElement.innerText = `Error generating legal response: ${errorMessage}`;
            return;
          }
          
          // Process successful response
          const data = await response.json();
          console.log("Response data:", data);
          
          if (data.choices && data.choices.length > 0) {
            const aiResponse = data.choices[0].message.content.trim();
            
            // Extract the draft email content from the response
            let displayResponse = aiResponse;
            
            // First try to get content after "DRAFT EMAIL:" marker
            if (aiResponse.includes("DRAFT EMAIL:")) {
              displayResponse = aiResponse.split("DRAFT EMAIL:")[1].trim();
            }
            
            // Also handle the case where response contains "### Draft Email:"
            if (displayResponse.includes("### Draft Email:")) {
              displayResponse = displayResponse.split("### Draft Email:")[1].trim();
            }
            
            // Remove any analysis section if present
            if (displayResponse.includes("### Analysis:")) {
              displayResponse = displayResponse.split("### Analysis:")[1].trim();
              // If the response contains both analysis and draft email sections
              if (displayResponse.includes("### Draft Email:")) {
                displayResponse = displayResponse.split("### Draft Email:")[1].trim();
              }
            }
            
            // Show the response and action buttons
            responseContainer.innerText = displayResponse;
            responseContainer.style.display = "block";
            copyButton.style.display = "inline-block";
            replyButton.style.display = "inline-block";
            statusElement.innerText = "Email response generated!";
          } else {
            statusElement.innerText = "No content in the response from Azure OpenAI.";
          }
        } catch (fetchError) {
          console.error("Fetch error:", fetchError);
          statusElement.innerText = `Network error: ${fetchError.message}`;
        }
      } catch (processingError) {
        console.error("Processing error:", processingError);
        statusElement.innerText = `Error processing email: ${processingError.message}`;
      }
    });
  } catch (error) {
    console.error("General error:", error);
    statusElement.innerText = `Error: ${error.message}`;
  }
}

/**
 * Creates a reply to the current email with the generated response
 */
function replyWithResponse() {
  const statusElement = document.getElementById("status");
  const responseContainer = document.getElementById("response-container");
  
  try {
    statusElement.innerText = "Creating reply all...";
    
    // Get the response text
    let responseText = responseContainer.innerText;
    if (!responseText) {
      statusElement.innerText = "No response generated yet.";
      return;
    }
    
    // Remove any Analysis or Draft Email markers if they weren't cleaned up earlier
    if (responseText.includes("### Analysis:")) {
      responseText = responseText.split("### Analysis:")[1].trim();
      if (responseText.includes("### Draft Email:")) {
        responseText = responseText.split("### Draft Email:")[1].trim();
      }
    } else if (responseText.includes("### Draft Email:")) {
      responseText = responseText.split("### Draft Email:")[1].trim();
    }
    
    // Get the current item (email)
    const item = Office.context.mailbox.item;
    
    if (!item) {
      statusElement.innerText = "No email selected";
      return;
    }

    // Format the response with proper HTML styling
    let formattedHtml = responseText;
    
    // Extract subject if it exists in the format **Subject: Some subject text**
    let subjectText = "";
    const subjectMatch = formattedHtml.match(/\*\*Subject:\s*(.*?)\*\*/i);
    if (subjectMatch && subjectMatch[1]) {
      subjectText = subjectMatch[1].trim();
      // Remove the subject line from the body but keep it in another variable
      formattedHtml = formattedHtml.replace(/\*\*Subject:\s*(.*?)\*\*\s*\n?/i, "");
    }
    
    // Convert markdown-style bold (**text**) to HTML bold
    formattedHtml = formattedHtml.replace(/\*\*(.*?)\*\*/g, "<strong>$1</strong>");
    
    // Convert URLs to hyperlinks (if not already in markdown format)
    formattedHtml = formattedHtml.replace(
      /(?<![\[\(])(https?:\/\/[^\s\)]+)(?![\]\)])/g, 
      '<a href="$1">$1</a>'
    );
    
    // Preserve markdown style links [text](url)
    formattedHtml = formattedHtml.replace(
      /\[([^\]]+)\]\(([^)]+)\)/g,
      '<a href="$2">$1</a>'
    );
    
    // Convert newlines to <br> tags
    formattedHtml = formattedHtml.replace(/\n/g, "<br>");
    
    // Use the subject from OpenAI response, or use the existing subject if none was provided
    const replySubject = subjectText || 
                        (item.subject.startsWith("Conference -") ? 
                          item.subject : 
                          "Conference - " + item.subject);
    
    // Add the subject as plain text at the top of the email body for easy copying
    let subjectHeader = "";
    if (subjectText) {
      subjectHeader = `${subjectText}<br><br>`;
    }
    
    // Wrap the entire content in a div with Arial font and size 10
    formattedHtml = `<div style="font-family: Arial, sans-serif; font-size: 10pt;">
      ${subjectHeader}
      ${formattedHtml}
    </div>`;
    
    // Create a Reply All with our HTML body
    Office.context.mailbox.item.displayReplyAllForm({
      htmlBody: formattedHtml,
      subject: replySubject // Note: Outlook may still add "Re: " prefix
    });
    
    statusElement.innerText = "Reply All created with formatted response. The subject line is at the top of the email for easy copying.";
  } catch (error) {
    console.error("Reply error:", error);
    statusElement.innerText = `Error creating reply: ${error.message}`;
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