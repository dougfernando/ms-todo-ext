/// <reference types="@raycast/api">

/* 🚧 🚧 🚧
 * This file is auto-generated from the extension's manifest.
 * Do not modify manually. Instead, update the `package.json` file.
 * 🚧 🚧 🚧 */

/* eslint-disable @typescript-eslint/ban-types */

type ExtensionPreferences = {
  /** Microsoft Graph API Token - Your personal access token for the Microsoft Graph API */
  "token": string
}

/** Preferences accessible in all the extension's commands */
declare type Preferences = ExtensionPreferences

declare namespace Preferences {
  /** Preferences accessible in the `list-todos` command */
  export type ListTodos = ExtensionPreferences & {}
  /** Preferences accessible in the `create-todo` command */
  export type CreateTodo = ExtensionPreferences & {}
}

declare namespace Arguments {
  /** Arguments passed to the `list-todos` command */
  export type ListTodos = {}
  /** Arguments passed to the `create-todo` command */
  export type CreateTodo = {}
}

