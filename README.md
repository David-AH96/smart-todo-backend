# Smart Todo - Backend Bridge

## 🎯 Purpose
This is the Node.js/Express backend for the Smart Todo application. It serves as a secure, local orchestration bridge between your React frontend and your communication tools. 

Because web browsers restrict direct access to local desktop apps and often block direct cross-origin API calls (CORS), this backend acts as the middleman to:
1. **Read Local Outlook Data:** It executes native macOS JXA (JavaScript for Automation) scripts to read emails and tasks directly from your running **Legacy Outlook for Mac** desktop app—without needing any cloud configuration or Azure OAuth tokens.
2. **Proxy Webex APIs:** It securely proxies requests to the Webex Cloud REST API, handling authentication via your Personal Access Token so your frontend doesn't have to expose it.

## 📋 Prerequisites
* **macOS** (Required for the Outlook AppleScript/JXA bridge to function).
* **Legacy Outlook for Mac** must be running in the background.
* **Node.js** installed on your machine.

## 🚀 Setup & Installation

**1. Install Dependencies**
Open your terminal, navigate to this backend folder, and run:
```bash
npm install express cors dotenv node-fetch

2. Secure Your Webex Token
To fetch your chat action items, the server needs your Webex Personal Access Token. To keep this secure and out of GitHub, we use a hidden .env file.

Step A: Go to developer.webex.com and log in with your Webex account.
Step B: Click on your profile avatar in the top right corner and copy your Personal Access Token.
Step C: In the root folder of this backend project (right next to server.js), create a new file named exactly .env.
Step D: Open the .env file and paste your token in this exact format:
WEBEX_TOKEN=your_copied_personal_access_token_here

Once your dependencies are installed and your .env file is ready, start the server:
node server.js
