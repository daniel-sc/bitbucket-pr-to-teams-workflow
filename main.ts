function getRequiredEnv(name: string): string {
  const value = Deno.env.get(name)?.trim();

  if (!value) {
    console.error(`Environment variable ${name} must be set.`);
    Deno.exit(1);
  }

  return value;
}

const TARGET_USERS = new Set(
  (Deno.env.get("TARGET_USERS") ?? Deno.env.get("TARGET_USER") ?? "")
    .split(",")
    .map((user) => user.trim().toLowerCase())
    .filter(Boolean),
);
const TEAMS_WEBHOOK_URL = getRequiredEnv("TEAMS_WEBHOOK_URL");

if (TARGET_USERS.size === 0) {
  console.error(
    "Environment variable TARGET_USERS (or TARGET_USER) must be set.",
  );
  Deno.exit(1);
}

function normalizeUser(value: unknown): string | null {
  if (typeof value !== "string") {
    return null;
  }

  const normalized = value.trim().toLowerCase();
  return normalized || null;
}

// Function to send a message to MS Teams
async function postToTeams(
  description: string,
  viewUrl: string,
): Promise<void> {
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
              wrap: true,
            },
          ],
          actions: [
            {
              type: "Action.OpenUrl",
              title: "Open PR",
              url: viewUrl,
            },
          ],
          $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
          version: "1.5",
        },
      },
    ],
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

Deno.serve(async (req) => {
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

  const prAuthor = normalizeUser(
    payload.pullRequest?.author?.user?.emailAddress,
  );
  const authorName = payload.pullRequest?.author?.user?.displayName;

  if (prAuthor && TARGET_USERS.has(prAuthor)) {
    const prTitle = payload.pullRequest?.title || "(no title)";
    const prLink = payload.pullRequest?.links?.self[0]?.href || "(no link)";

    const description =
      `A new pull request was created by ${authorName}:\n**${prTitle}**`;

    await postToTeams(description, prLink);
  } else {
    console.log(`Skipping PR from non-target author: ${prAuthor ?? "unknown"}`);
  }

  return new Response("OK", { status: 200 });
});
