import { CallClient, CallAgent } from "@azure/communication-calling";
import { AzureCommunicationTokenCredential } from "@azure/communication-common";
import { CommunicationIdentityClient } from "@azure/communication-identity";
import { ChatClient } from "@azure/communication-chat";

// Variables
let userId = "";
let tokenString = "";
let call;
let callAgent;
let chatClient;
let chatThreadClient;

// Config
// TODO: Get from database
const chatThreadId = "";
const joinMeetingUrl = "";
const endpointUrl = "";

// Secret to get from KV
const connectionString = "";

// 1. Create client and user on ACS
const identityClient = new CommunicationIdentityClient(connectionString);
let identityResponse = await identityClient.createUser();
userId = identityResponse.communicationUserId;
console.log(
  `\nCreated an identity with ID: ${identityResponse.communicationUserId}`
);

// 2. Define ACS scopes and get ACS token
let tokenResponse = await identityClient.getToken(identityResponse, [
  "voip",
  "chat",
]);

const { token, expiresOn } = tokenResponse;
tokenString = token;
console.log(`\nIssued an access token that expires at: ${expiresOn}`);
console.log(token);

// 3. Create call/chat client
const callClient = new CallClient();
const tokenCredential = new AzureCommunicationTokenCredential(token);
callAgent = await callClient.createCallAgent(tokenCredential);
chatClient = new ChatClient(
  endpointUrl,
  new AzureCommunicationTokenCredential(token)
);

// 4. Join meeting
call = callAgent.join(
  {
    meetingLink: joinMeetingUrl,
  },
  {}
);
console.log(call);

// 5. Join chat
await chatClient.startRealtimeNotifications();

// subscribe to new message notifications
chatClient.on("chatMessageReceived", (e) => {
  console.log(`Notification chatMessageReceived! ${e.message}`);

  // check whether the notification is intended for the current thread
  if (chatThreadId != e.threadId) {
    return;
  }

  if (e.sender.communicationUserId != userId) {
    // renderReceivedMessage(e);
  } else {
    // renderSentMessage(e.message);
  }
});

chatThreadClient = await chatClient.getChatThreadClient(chatThreadId);
console.log("Azure Communication Chat client created!");

// 6. Send chat message
let message = "Hello";

let sendMessageRequest = {
  content: message,
};
let sendMessageOptions = {
  senderDisplayName: "Jack",
};
let sendChatMessageResult = await chatThreadClient.sendMessage(
  sendMessageRequest,
  sendMessageOptions
);
let messageId = sendChatMessageResult.id;

console.log(`Message sent!, message id:${messageId}`);
