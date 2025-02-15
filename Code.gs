/**
 * This function builds the add-on’s UI when an email is opened.
 * It is triggered via the contextual trigger in the manifest.
 */
/**
 * Combined buildAddOn function that shows the email subject, classification, summary,
 * and provides quick feedback buttons and a detailed feedback text input.
 */
function buildAddOn(e) {
  // Fallback if the event doesn't include email metadata.
  if (!e || !e.messageMetadata) {
    return CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle("No Email Context"))
      .addSection(
        CardService.newCardSection().addWidget(
          CardService.newTextParagraph().setText("Please open an email to use this add-on.")
        )
      )
      .build();
  }
  
  var messageId = e.messageMetadata.messageId;
  var message = GmailApp.getMessageById(messageId);
  var subject = message.getSubject();
  var snippet = message.getPlainBody().substring(0, 200);
  
  // Get the classification and summary from OpenAI.
  var classification = classifyEmail(subject, snippet);
  var summary = summarizeEmail(subject, snippet);
  
  // Build the add-on card.
  var card = CardService.newCardBuilder();
  var section = CardService.newCardSection()
      .addWidget(CardService.newTextParagraph().setText("<b>Subject:</b> " + subject))
      .addWidget(CardService.newTextParagraph().setText("<b>Classification:</b> " + classification))
      .addWidget(CardService.newTextParagraph().setText("<b>Summary:</b> " + summary));
  
  // Create quick feedback actions.
  var correctAction = CardService.newAction()
    .setFunctionName("handleFeedback")
    .setParameters({ feedback: "correct", messageId: messageId });
  var incorrectAction = CardService.newAction()
    .setFunctionName("handleFeedback")
    .setParameters({ feedback: "incorrect", messageId: messageId });
  section.addWidget(
    CardService.newButtonSet()
      .addButton(CardService.newTextButton().setText("👍 Correct").setOnClickAction(correctAction))
      .addButton(CardService.newTextButton().setText("👎 Incorrect").setOnClickAction(incorrectAction))
  );
  
  // Detailed feedback text input.
  section.addWidget(
    CardService.newTextInput()
      .setFieldName("detailedFeedback")
      .setTitle("Detailed Feedback (Optional)")
      .setHint("Enter additional feedback...")
  );
  
  // Submit detailed feedback button.
  var submitAction = CardService.newAction()
      .setFunctionName("handleDetailedFeedback")
      .setParameters({ messageId: messageId });
  section.addWidget(
    CardService.newTextButton().setText("Submit Detailed Feedback").setOnClickAction(submitAction)
  );
  
  card.addSection(section);
  return card.build();
}

/**
 * Handles detailed feedback submission from the user.
 */
function handleDetailedFeedback(e) {
  var formInputs = e.commonEventObject.formInputs;
  var messageId = e.parameters.messageId;
  var detailedFeedback = formInputs.detailedFeedback.stringInputs.value[0];
  
  Logger.log("Detailed feedback for message " + messageId + ": " + detailedFeedback);
  
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText("Detailed feedback submitted. Thank you!"))
    .build();
}

/**
 * Optional: Show a configuration card to let users define custom labels.
 */
function showConfigCard() {
  var userProperties = PropertiesService.getUserProperties();
  var savedLabels = userProperties.getProperty("customLabels") || "Important Information,Unimportant,Requires Action,Requires Reply";
  
  var card = CardService.newCardBuilder();
  var section = CardService.newCardSection()
      .addWidget(
        CardService.newTextInput()
          .setFieldName("labelsInput")
          .setTitle("Custom Labels")
          .setHint("Enter labels separated by commas")
          .setValue(savedLabels)
      );
  
  var saveAction = CardService.newAction().setFunctionName("saveCustomLabels");
  section.addWidget(
    CardService.newTextButton()
      .setText("Save Labels")
      .setOnClickAction(saveAction)
  );
  
  card.addSection(section);
  return card.build();
}

/**
 * Saves the custom labels input by the user.
 */
function saveCustomLabels(e) {
  var formInputs = e.commonEventObject.formInputs;
  var labelsInput = formInputs.labelsInput.stringInputs.value[0];
  PropertiesService.getUserProperties().setProperty("customLabels", labelsInput);
  
  return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText("Custom labels saved!"))
      .build();
}



/**
 * Calls the OpenAI API to classify an email.
 * (You must replace YOUR_OPENAI_API_KEY with your actual API key.)
 */
function classifyEmail(subject, snippet) {
  var content = "Subject: " + subject + "\nSnippet: " + snippet;
  var prompt = "Classify the following email into one of these categories: Important Information, Unimportant, Requires Action, Requires Reply.\n\nEmail Content:\n" + content + "\n\nCategory:";

  var apiKey = "API_KEY"; // Replace with your actual API key.
  var url = "https://api.openai.com/v1/chat/completions";
  
  var payload = {
    "model": "gpt-4o",  // Use "gpt-4" if you have access.
    "messages": [
      {"role": "system", "content": "You are a helpful email categorizer."},
      {"role": "user", "content": prompt}
    ],
    "max_tokens": 10,
    "temperature": 0.3
  };

  var options = {
    "method": "post",
    "contentType": "application/json",
    "headers": {
      "Authorization": "Bearer " + apiKey
    },
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    Logger.log("ChatGPT response code: " + response.getResponseCode());
    Logger.log("ChatGPT response text: " + response.getContentText());
    var result = JSON.parse(response.getContentText());
    if (result && result.choices && result.choices.length > 0 && result.choices[0].message) {
      return result.choices[0].message.content.trim();
    }
  } catch (err) {
    Logger.log("Error calling ChatGPT API: " + err);
  }
  return "Unknown";
}



/**
 * Handles quick feedback (Correct/Incorrect) when a user clicks one of the buttons.
 */
function handleFeedback(e) {
  var params = e.parameters;
  var feedback = params.feedback;
  var messageId = params.messageId;
  
  // Log the feedback; in a real app you might store this data.
  Logger.log("Feedback for message " + messageId + ": " + feedback);
  
  // Build a simple confirmation card.
  var card = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle("Feedback Submitted"))
      .addSection(CardService.newCardSection().addWidget(
          CardService.newTextParagraph().setText("Thank you for your feedback!")
      ));
  return card.build();
}

/**
 * Handles detailed feedback submission from the text input.
 */
function handleDetailedFeedback(e) {
  var formInputs = e.commonEventObject.formInputs;
  var messageId = e.parameters.messageId;
  var detailedFeedback = formInputs.detailedFeedback.stringInputs.value[0];
  
  Logger.log("Detailed feedback for message " + messageId + ": " + detailedFeedback);
  
  // Build a confirmation card.
  var card = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle("Feedback Submitted"))
      .addSection(CardService.newCardSection().addWidget(
          CardService.newTextParagraph().setText("Thank you for your detailed feedback!")
      ));
  return card.build();
}

/**
 * This function processes the inbox, classifies emails, and applies labels.
 */
function classifyAndLabelEmails() {
  // Fetch a batch of inbox threads (adjust count as needed)
  var threads = GmailApp.getInboxThreads(0, 50);
  
  // Process each thread.
  for (var i = 0; i < threads.length; i++) {
    var thread = threads[i];
    // We'll process only the first message in the thread for simplicity.
    var message = thread.getMessages()[0];
    
    // Skip if the thread already has a classification label.
    var existingLabels = thread.getLabels().map(function(label) {
      return label.getName();
    });
    
    // (Optional) Define your set of classification labels
    var possibleLabels = ["Important Information", "Unimportant", "Requires Action", "Requires Reply"];
    var alreadyLabeled = possibleLabels.some(function(label) {
      return existingLabels.indexOf(label) !== -1;
    });
    
    if (alreadyLabeled) {
      continue;  // Skip threads that have already been labeled.
    }
    
    // Get subject and a snippet (first 200 characters)
    var subject = message.getSubject();
    var snippet = message.getPlainBody().substring(0, 200);
    
    // Get the classification from OpenAI.
    var category = classifyEmail(subject, snippet);
    
    // Create or get the label for this category.
    var label = GmailApp.getUserLabelByName(category);
    if (!label) {
      label = GmailApp.createLabel(category);
    }
    
    // Apply the label to the entire thread.
    thread.addLabel(label);
    
    // Log the operation for debugging.
    Logger.log("Thread labeled: " + subject + " => " + category);
  }
}

function summarizeEmail(subject, snippet) {
  // Build the prompt for summarization
  var emailContent = "Subject: " + subject + "\n\nContent: " + snippet;
  var prompt = "Please provide a concise summary (a few sentences) of the following email:\n\n" + emailContent;
  
  var apiKey = "API_KEY"; // Replace with your API key.
  var url = "https://api.openai.com/v1/chat/completions";
  
  var payload = {
    "model": "gpt-4o", // or use "gpt-4" if you have access.
    "messages": [
      {"role": "system", "content": "You are a helpful email summarizer."},
      {"role": "user", "content": prompt}
    ],
    "max_tokens": 60,
    "temperature": 0.5
  };
  
  var options = {
    "method": "post",
    "contentType": "application/json",
    "headers": {
      "Authorization": "Bearer " + apiKey
    },
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };
  
  try {
    var response = UrlFetchApp.fetch(url, options);
    Logger.log("Summarizer response code: " + response.getResponseCode());
    Logger.log("Summarizer response text: " + response.getContentText());
    var result = JSON.parse(response.getContentText());
    if (result && result.choices && result.choices.length > 0 && result.choices[0].message) {
      return result.choices[0].message.content.trim();
    }
  } catch (err) {
    Logger.log("Error calling summarizer: " + err);
  }
  return "Summary not available.";
}
