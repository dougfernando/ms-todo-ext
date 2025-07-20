/// <reference types="@raycast/api">

/* 🚧 🚧 🚧
 * This file is auto-generated from the extension's manifest.
 * Do not modify manually. Instead, update the `package.json` file.
 * 🚧 🚧 🚧 */

/* eslint-disable @typescript-eslint/ban-types */

type ExtensionPreferences = {
  /** Client ID - The Client ID of your application registration in Azure. */
  "clientId": string
}

/** Preferences accessible in all the extension's commands */
declare type Preferences = ExtensionPreferences

declare namespace Preferences {
  /** Preferences accessible in the `list-todos` command */
  export type ListTodos = ExtensionPreferences & {}
  /** Preferences accessible in the `create-todo` command */
  export type CreateTodo = ExtensionPreferences & {}
  /** Preferences accessible in the `list-tasks-by-list` command */
  export type ListTasksByList = ExtensionPreferences & {}
}

declare namespace Arguments {
  /** Arguments passed to the `list-todos` command */
  export type ListTodos = {}
  /** Arguments passed to the `create-todo` command */
  export type CreateTodo = {}
  /** Arguments passed to the `list-tasks-by-list` command */
  export type ListTasksByList = {}
}

