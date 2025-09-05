# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Raycast extension for Microsoft To Do that allows users to view, create, and complete tasks from within Raycast. It's built as a TypeScript React application using the Raycast API and Microsoft Graph API.

## Key Commands

- `npm ci` - Install dependencies (recommended over npm install)
- `npm run dev` - Start development mode (adds extension to Raycast)  
- `npm run build` - Build the extension for distribution
- `npm run lint` - Run ESLint to check for code issues
- `npm run fix-lint` - Automatically fix linting issues
- `ray build -e dist` - Build extension (same as npm run build)
- `ray develop` - Start development mode (same as npm run dev)
- `ray publish` - Publish extension to Raycast store

## Architecture

The extension consists of three main commands defined in package.json:

1. **list-todos** (`src/list-todos.tsx`) - Main command that displays all tasks grouped by lists
2. **create-todo** (`src/create-todo.tsx`) - Form for creating new tasks with due dates and importance levels
3. **list-tasks-by-list** (`src/list-tasks-by-list.tsx`) - Two-level navigation to browse lists then tasks

### Authentication Flow

- Uses Microsoft OAuth2 PKCE flow via `src/auth.ts`
- Requires Azure app registration with Client ID configured in preferences
- Handles token refresh automatically
- Scopes: `Tasks.ReadWrite offline_access`

### API Integration

- All commands use Microsoft Graph API (`https://graph.microsoft.com/v1.0/me/todo/`)
- Rate limiting protection in `list-todos.tsx` with retry logic and throttling handling
- Filters out "Flagged Emails" system list
- Only shows incomplete tasks by default (`$filter=status ne 'completed'`)

### State Management

Each command manages its own local state using React hooks:
- Loading states with toast notifications
- Optimistic UI updates for task completion
- Error handling with user-friendly messages

## Setup Requirements

1. Node.js installation required
2. Azure Portal app registration needed for Client ID
3. Raycast installed and configured
4. Client ID must be added to extension preferences after installation

## Development Notes

- Uses `p-limit` for API concurrency control (set to 3 concurrent requests)
- Implements sleep/retry mechanisms for Microsoft Graph API throttling
- Form validation handled by Raycast Form components
- All API calls include proper error handling and user feedback