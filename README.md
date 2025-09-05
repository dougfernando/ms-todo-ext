# Microsoft To Do Extension for Raycast

A comprehensive Raycast extension that seamlessly integrates Microsoft To Do into your workflow. Manage your tasks efficiently without leaving Raycast.

## ✨ Features

- **📋 View All Tasks**: Browse all your tasks organized by lists with due date indicators
- **✅ Quick Completion**: Mark tasks as complete with a single action
- **➕ Create Tasks**: Add new tasks with titles, notes, due dates, and importance levels
- **📂 List Navigation**: Browse tasks by specific lists for better organization
- **📅 Due Date Display**: See due dates at a glance ("Today", "Tomorrow", or formatted dates)
- **🔄 Real-time Sync**: Automatic synchronization with Microsoft To Do
- **🛡️ Secure Authentication**: OAuth 2.0 integration with Microsoft accounts

## 🚀 Installation

### Prerequisites
- [Node.js](https://nodejs.org/) (install via `winget install -e --id OpenJS.NodeJS`)
- [Raycast](https://raycast.com/)
- Microsoft account with To Do access

### Setup Steps
1. Clone this repository
2. Run `npm ci` to install dependencies
3. Run `npm run dev` to add the extension to Raycast
4. Configure your Azure app registration (see Authentication Setup below)

## 🔐 Authentication Setup

This extension requires an Azure app registration for Microsoft Graph API access:

1. Go to [Azure Portal](https://portal.azure.com/)
2. Navigate to "App registrations" → "New registration"
3. Set up your app with these settings:
   - **Name**: Microsoft To Do Raycast Extension
   - **Supported account types**: Accounts in any organizational directory and personal Microsoft accounts
   - **Redirect URI**: Add the Raycast OAuth redirect URI
4. Copy the **Application (client) ID**
5. Add the Client ID to the extension preferences in Raycast

### Required API Permissions
- `Tasks.ReadWrite` - Read and write access to user tasks
- `offline_access` - Maintain access to data when user is offline

## 📱 Available Commands

### List To-Dos
View all your tasks grouped by lists, with due date indicators and quick completion actions.

### Create To-Do
Create new tasks with:
- Title and detailed notes
- Due date selection
- Importance levels (Low, Normal, High)
- List assignment
- Optional completion status

### List Tasks by List
Browse your task lists first, then drill down to see tasks within specific lists.

## 🛠️ Development

```bash
# Install dependencies
npm ci

# Start development mode
npm run dev

# Build for production
npm run build

# Lint code
npm run lint

# Auto-fix linting issues
npm run fix-lint
```

## 📝 Notes

- Only incomplete tasks are displayed by default
- The "Flagged Emails" system list is automatically filtered out
- Tasks are updated optimistically for better user experience
- Built-in rate limiting and retry logic for Microsoft Graph API calls
- Due dates are displayed with smart formatting (Today/Tomorrow/Date)