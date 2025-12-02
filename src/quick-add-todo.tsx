// src/quick-add-todo.tsx
import {
    ActionPanel,
    Form,
    Action,
    showToast,
    Toast,
    closeMainWindow,
    getSelectedText,
    getPreferenceValues,
} from "@raycast/api"
import { useEffect, useState } from "react"
import fetch from "node-fetch"
import { authorize, getAccessToken } from "./auth"

// Interfaces
interface TaskList {
    id: string
    displayName: string
}

interface QuickTaskForm {
    title: string
    taskList: string
}

interface TaskResponse {
    id: string
}

interface Preferences {
    defaultList?: string
}

// API Request to create todo
async function createQuickTodo(task: QuickTaskForm): Promise<TaskResponse> {
    const token = await getAccessToken()
    const body = {
        title: task.title,
        importance: "normal",
    }

    const response = await fetch(`https://graph.microsoft.com/v1.0/me/todo/lists/${task.taskList}/tasks`, {
        method: "POST",
        headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
        body: JSON.stringify(body),
    })

    if (!response.ok) {
        const errorData: any = await response.json()
        throw new Error(errorData.error?.message || "Failed to create task")
    }

    return (await response.json()) as TaskResponse
}

// API Request to fetch task lists
async function fetchTaskLists(): Promise<TaskList[]> {
    const token = await getAccessToken()
    try {
        const response = await fetch("https://graph.microsoft.com/v1.0/me/todo/lists", {
            headers: { Authorization: `Bearer ${token}` },
        })
        const data: any = await response.json()

        if (data.error) {
            throw new Error(data.error.message)
        }
        return data.value.filter((taskList: TaskList) => taskList.displayName !== "Flagged Emails")
    } catch (error) {
        const message = error instanceof Error ? error.message : "An unknown error occurred"
        throw new Error(message)
    }
}

export default function QuickAddTodoCommand() {
    const [taskLists, setTaskLists] = useState<TaskList[]>([])
    const [isLoading, setIsLoading] = useState(true)
    const [selectedText, setSelectedText] = useState("")
    const preferences = getPreferenceValues<Preferences>()

    useEffect(() => {
        async function initializeCommand() {
            await authorize()
            setIsLoading(true)

            try {
                // Try to get selected text from other applications
                const text = await getSelectedText()
                if (text) {
                    setSelectedText(text)
                }
            } catch (error) {
                // No selected text is fine, just continue
            }

            try {
                const taskListsResponse = await fetchTaskLists()
                setTaskLists(taskListsResponse)
            } catch (error) {
                await showToast(Toast.Style.Failure, "Error", "Could not load task lists")
            }

            setIsLoading(false)
        }
        initializeCommand()
    }, [])

    async function handleSubmit(values: QuickTaskForm) {
        const toast = await showToast({ style: Toast.Style.Animated, title: "Creating task..." })

        try {
            await createQuickTodo(values)

            toast.style = Toast.Style.Success
            toast.title = "Task Created!"

            await closeMainWindow({ clearRootSearch: true })
        } catch (error) {
            toast.style = Toast.Style.Failure
            toast.title = "Error"
            toast.message = error instanceof Error ? error.message : "Could not create task"
        }
    }

    // Find default list or use first available list
    const defaultListId =
        preferences.defaultList ||
        (taskLists.length > 0
            ? taskLists.find(
                  list =>
                      list.displayName.toLowerCase().includes("task") ||
                      list.displayName.toLowerCase().includes("todo"),
              )?.id || taskLists[0].id
            : undefined)

    return (
        <Form
            isLoading={isLoading}
            actions={
                <ActionPanel>
                    <Action.SubmitForm title="Create Task" icon="⚡" onSubmit={handleSubmit} />
                </ActionPanel>
            }
        >
            <Form.TextField
                id="title"
                title="Task Title"
                placeholder="What needs to be done?"
                defaultValue={selectedText}
                autoFocus
            />
            <Form.Dropdown id="taskList" title="List" defaultValue={defaultListId}>
                {taskLists.map(list => (
                    <Form.Dropdown.Item key={list.id} value={list.id} title={list.displayName} />
                ))}
            </Form.Dropdown>
        </Form>
    )
}
