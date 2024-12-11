// Import necessary modules
import { serve } from "https://deno.land/std@0.203.0/http/server.ts";

// Load environment variables
const TARGET_USER = Deno.env.get("TARGET_USER");
const TEAMS_WEBHOOK_URL = Deno.env.get("TEAMS_WEBHOOK_URL");

if (!TARGET_USER || !TEAMS_WEBHOOK_URL) {
  console.error("Environment variables TARGET_USER and TEAMS_WEBHOOK_URL must be set.");
  Deno.exit(1);
}

// Function to send a message to MS Teams
async function postToTeams(description: string, viewUrl: string): Promise<void> {
  const adaptiveCard = {
    type: "message",
    attachments: [
      {
        contentType: "application/vnd.microsoft.card.adaptive",
        content: {
          type: "AdaptiveCard",
          body: [
            {
              type: "TextBlock",
              text: description,
              wrap: true
            }
          ],
          actions: [
            {
              type: "Action.OpenUrl",
              title: "Open PR",
              url: viewUrl
            }
          ],
          $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
          version: "1.5"
        }
      }
    ]
    
  };

  const response = await fetch(TEAMS_WEBHOOK_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(adaptiveCard),
  });

  if (!response.ok) {
    console.error(`Failed to post to Teams: ${response.statusText}`);
  } else {
    console.log(`Send post to Teams: ${response.statusText}`);
  }
}

// Start the server
serve(async (req) => {
  if (req.method !== "POST") {
    return new Response("Method not allowed", { status: 405 });
  }

  let payload;
  try {
    payload = await req.json();
  } catch (err) {
    console.error("Failed to parse JSON payload:", err);
    return new Response("Invalid JSON", { status: 400 });
  }

  // Check if the webhook event is a pull request creation
  if (payload.eventKey !== "pr:opened") {
    return new Response("Event not handled", { status: 200 });
  }

  const prAuthor = payload.pullRequest?.author?.user?.emailAddress;
  const authorName = payload.pullRequest?.author?.user?.displayName;

  if (prAuthor === TARGET_USER) {
    const prTitle = payload.pullRequest?.title || "(no title)";
    const prLink = payload.pullRequest?.links?.self[0]?.href || "(no link)";

    const description = `A new pull request was created by ${authorName}:\n**${prTitle}**`;

    await postToTeams(description, prLink);
  } else {
    console.log(`skipping PR from different author: ${prAuthor}`)
  }

  return new Response("OK", { status: 200 });
});
