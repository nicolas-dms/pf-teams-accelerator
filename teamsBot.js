const { TeamsActivityHandler, TurnContext } = require("botbuilder");

class TeamsBot extends TeamsActivityHandler {
  constructor() {
    super();

    this.onMessage(async (context, next) => {
      
      console.log("Running with Message Activity.");
      const removedMentionText = TurnContext.removeRecipientMention(context.activity);
      const txt = removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim();
      await this.callExternalStreamingAPI(txt,context);
      // By calling next() you ensure that the next BotHandler is run.
      await next();
    });

    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded;
      for (let cnt = 0; cnt < membersAdded.length; cnt++) {
        if (membersAdded[cnt].id) {
          await context.sendActivity(
            `Hi there! I'm a Teams bot that will echo what you said to me.`
          );
          break;
        }
      }
      await next();
    });
  }

  async callExternalStreamingAPI(inputText,context) {
    const requestBody = JSON.stringify({ 
      question: inputText,
      chat_history: [] 
    });
    
    // Set up the request headers, including the text/event-stream header
    const requestHeaders = new Headers({
      "Accept": "text/event-stream",  // Required for event streams
      "Content-Type": "application/json"
    });

    const apiKey = "UTDhglCp82tc0T3kRh4iKPdePcKPQ6R9";

    if (!apiKey) {
        throw new Error("A key should be provided to invoke the endpoint");
    }
    requestHeaders.append("Authorization", "Bearer " + apiKey);

    requestHeaders.append("azureml-model-deployment", "sg-chat-demo-buafz-2");

    const url = "https://sg-chat-demo-buafz.francecentral.inference.ml.azure.com/score";

    try {
      const response = await fetch(url, {
          method: "POST",
          body: requestBody,
          headers: requestHeaders,
      });

      if (!response.ok) {
          const responseBody = await response.text();
          console.debug('Response Body:', responseBody);
          throw new Error("Request failed with status code " + response.status);
      }

      const reader = response.body.getReader();
      const decoder = new TextDecoder('utf-8');
      let sentenceBuffer = '';

      function extractAndConcatAnswers(chunk) {
        // Split the chunk into parts based on "data:" prefix
        const parts = chunk.split(/data:\s*/);
    
        let concatenatedAnswer = '';
    
        // Iterate through each part and process if it's a valid JSON
        parts.forEach(part => {
            // check if the part contains a valid JSON object
            if (part) {
                try {
                    // Parse the JSON and extract the 'answer' field
                    const json = JSON.parse(part);
                    if (json.answer) {
                        concatenatedAnswer += json.answer;  // Concatenate the answer field
                    }
                } catch (error) {
                    console.error("Failed to parse JSON:", error);
                }
            }
        });
        return concatenatedAnswer.trim();  // Return the concatenated answer
      }

      // Initialize a buffer to store sentences
      let messageBuffer = '';

      while (true) {
        const { done, value } = await reader.read();
        if (done) break;

        // Decode the chunk into a string
        const chunk = decoder.decode(value, { stream: true });

        // Remove the 'data: ' prefix from the response chunk
        try {
          console.log('Received chunk:', chunk);
          const responseStr = extractAndConcatAnswers(chunk)
          const cleanedResponseStr = responseStr.replace(/#+/g, '').replace(':',':\n');  // Remove any hash characters
          console.log('cleaned chunk:', cleanedResponseStr);
          sentenceBuffer += cleanedResponseStr + " ";
        } catch (error) {
            console.error('Failed to parse JSON:', error);
            console.error('Error Chunk:', chunk);
            continue;
        }

        // Split the received data into sentences
        const sentences = sentenceBuffer.split('.');

        for (let i = 0; i < sentences.length - 1; i++) {
            // Accumulate the sentence into the messageBuffer
            messageBuffer += sentences[i].trim() + '.\n';

            // Check if the messageBuffer length is between 100 and 200 characters
            if (messageBuffer.length >= 20 && messageBuffer.length <= 50) {
                // Send the accumulated message as an activity
                await context.sendActivity(messageBuffer.trim());
                // Reset the buffer after sending
                messageBuffer = '';
                this.delay(500)
            }
        }
        // Keep the incomplete sentence for the next chunk
        sentenceBuffer = sentences[sentences.length - 1];
      }

      // After all chunks are processed, send any remaining text if it is the last buffer
      if (sentenceBuffer.length > 0 || messageBuffer.length > 0) {
        // If the remaining text is smaller than 100 characters, send it as the final activity
        await context.sendActivity((messageBuffer + sentenceBuffer).trim());
      }
    } catch (error) {
      console.error("Error during API call:", error);
      await context.sendActivity("Sorry, something went wrong with the API call.");
    }
  }

  // Helper function to introduce a delay
  delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }
}

module.exports.TeamsBot = TeamsBot;
