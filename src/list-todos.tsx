// src/list-todos.tsx
import { ActionPanel, Action, List, showToast, Toast, Icon } from "@raycast/api";
import { useEffect, useState } from "react";
import fetch, { Response } from "node-fetch"; // Import Response type
import pLimit from "p-limit";
import { authorize, getAccessToken } from "./auth";

// Interfaces
interface TaskList {
    id: string;
    displayName: string;
}
interface Todo {
    id: string;
    title: string;
    status: string;
    dueDateTime?: {
        dateTime: string;
        timeZone: string;
    };
}
interface GroupedTodos {
    list: TaskList;
    todos: Todo[];
}

// NEW: A simple helper function to pause execution for a given number of milliseconds.
const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

// Helper function to format due date
function formatDueDate(dueDateTime?: { dateTime: string; timeZone: string }): string | undefined {
    if (!dueDateTime) return undefined;
    
    try {
        const date = new Date(dueDateTime.dateTime);
        const today = new Date();
        const tomorrow = new Date(today);
        tomorrow.setDate(today.getDate() + 1);
        
        // Reset time to compare only dates
        const dueDateOnly = new Date(date.getFullYear(), date.getMonth(), date.getDate());
        const todayOnly = new Date(today.getFullYear(), today.getMonth(), today.getDate());
        const tomorrowOnly = new Date(tomorrow.getFullYear(), tomorrow.getMonth(), tomorrow.getDate());
        
        if (dueDateOnly.getTime() === todayOnly.getTime()) {
            return "Today";
        } else if (dueDateOnly.getTime() === tomorrowOnly.getTime()) {
            return "Tomorrow";
        } else {
            return date.toLocaleDateString(undefined, { 
                month: 'short', 
                day: 'numeric',
                year: date.getFullYear() !== today.getFullYear() ? 'numeric' : undefined
            });
        }
    } catch (error) {
        return undefined;
    }
}

// API Call to fetch To-Do lists
async function fetchTaskLists(): Promise<TaskList[]> {
    const token = await getAccessToken();
    try {
        const response = await fetch("https://graph.microsoft.com/v1.0/me/todo/lists", {
            headers: { Authorization: `Bearer ${token}` },
        });
        if (!response.ok) {
            const data = (await response.json()) as { error?: { message: string } };
            throw new Error(data.error?.message || "Failed to fetch task lists");
        }
        const data: any = await response.json();
        return data.value.filter((list: any) => list.displayName !== "Flagged Emails");
    } catch (error) {
        throw new Error(error instanceof Error ? error.message : "Could not fetch task lists");
    }
}

// UPDATED: fetchTodosForList now includes a retry mechanism for throttling errors
async function fetchTodosForList(listId: string): Promise<Todo[]> {
    const token = await getAccessToken();
    const maxRetries = 3;
    let attempt = 0;

    while (attempt < maxRetries) {
        try {
            const response: Response = await fetch(
                `https://graph.microsoft.com/v1.0/me/todo/lists/${listId}/tasks?$filter=status ne 'completed'`,
                {
                    headers: { Authorization: `Bearer ${token}` },
                },
            );

            // If throttled (HTTP 429), wait and retry
            if (response.status === 429) {
                // The API often provides a 'Retry-After' header with the number of seconds to wait.
                const retryAfterSeconds = parseInt(response.headers.get("Retry-After") || "5", 10);
                console.warn(`Throttled on list ${listId}. Retrying after ${retryAfterSeconds} seconds...`);
                await sleep(retryAfterSeconds * 1000);
                attempt++;
                continue; // Try the request again
            }

            if (!response.ok) {
                const data = (await response.json()) as { error?: { message: string } };
                throw new Error(data.error?.message || `HTTP error ${response.status}`);
            }

            const data: any = await response.json();
            return data.value; // Success! Exit the loop.
        } catch (error) {
            console.error(`Failed to fetch todos for list ${listId} on attempt ${attempt + 1}:`, error);
            attempt++;
            if (attempt >= maxRetries) {
                // If all retries fail, give up on this list and return empty.
                return [];
            }
            // Wait for an increasing amount of time before the next retry
            await sleep(2000 * attempt);
        }
    }

    return []; // Return empty if all retries fail
}

// API call to mark a task as complete
async function markTaskAsCompleteAPI(listId: string, taskId: string) {
    const token = await getAccessToken();
    return fetch(`https://graph.microsoft.com/v1.0/me/todo/lists/${listId}/tasks/${taskId}`, {
        method: "PATCH",
        headers: {
            Authorization: `Bearer ${token}`,
            "Content-Type": "application/json",
        },
        body: JSON.stringify({ status: "completed" }),
    });
}

export default function ListTodosCommand() {
    const [groupedTodos, setGroupedTodos] = useState<GroupedTodos[]>([]);
    const [isLoading, setIsLoading] = useState(true);

    async function loadTodos() {
        await authorize();
        setIsLoading(true);
        const toast = await showToast({ style: Toast.Style.Animated, title: "Loading tasks..." });
        try {
            // Lowering the concurrency limit to be safer
            const limit = pLimit(3);

            const taskLists = await fetchTaskLists();

            const promises = taskLists.map((list) =>
                limit(() => fetchTodosForList(list.id)).then((todos) => ({ list, todos })),
            );

            const results = await Promise.all(promises);
            const allGroupedTodos = results.filter((group) => group.todos.length > 0);
            setGroupedTodos(allGroupedTodos);

            toast.style = Toast.Style.Success;
            toast.title = "Tasks Loaded";
        } catch (error) {
            toast.style = Toast.Style.Failure;
            toast.title = "Error";
            toast.message = error instanceof Error ? error.message : "Could not load tasks";
        } finally {
            setIsLoading(false);
        }
    }

    useEffect(() => {
        loadTodos();
    }, []);

    async function handleMarkAsComplete(listId: string, taskId: string) {
        const originalTodos = [...groupedTodos];
        const newGroupedTodos = groupedTodos
            .map((group) => {
                if (group.list.id === listId) {
                    return { ...group, todos: group.todos.filter((t) => t.id !== taskId) };
                }
                return group;
            })
            .filter((group) => group.todos.length > 0);
        setGroupedTodos(newGroupedTodos);

        try {
            const response = await markTaskAsCompleteAPI(listId, taskId);
            if (!response.ok) {
                setGroupedTodos(originalTodos);
                const errorData: any = await response.json();
                await showToast(Toast.Style.Failure, "Failed to Complete Task", errorData.error?.message);
            } else {
                await showToast(Toast.Style.Success, "Task Completed!");
            }
        } catch (error) {
            setGroupedTodos(originalTodos);
            const message = error instanceof Error ? error.message : "An unknown error occurred";
            await showToast(Toast.Style.Failure, "Error", message);
        }
    }

    return (
        <List isLoading={isLoading} searchBarPlaceholder="Filter your to-dos...">
            {groupedTodos.length === 0 && !isLoading ? (
                <List.EmptyView title="No To-Dos Found" description="You're all caught up!" icon={Icon.Checkmark} />
            ) : (
                groupedTodos.map((group) => (
                    <List.Section key={group.list.id} title={group.list.displayName} subtitle={`${group.todos.length}`}>
                        {group.todos.map((todo) => (
                            <List.Item
                                key={todo.id}
                                title={todo.title}
                                icon={Icon.Circle}
                                accessories={formatDueDate(todo.dueDateTime) ? [{ text: formatDueDate(todo.dueDateTime) }] : undefined}
                                actions={
                                    <ActionPanel>
                                        <Action
                                            title="Mark as Complete"
                                            icon={Icon.CheckCircle}
                                            onAction={() => handleMarkAsComplete(group.list.id, todo.id)}
                                        />
                                        <Action.OpenInBrowser title="Open in To Do" url="https://to-do.live.com" />
                                        <Action
                                            title="Reload"
                                            icon={Icon.Repeat}
                                            onAction={loadTodos}
                                            shortcut={{ modifiers: ["cmd"], key: "r" }}
                                        />
                                    </ActionPanel>
                                }
                            />
                        ))}
                    </List.Section>
                ))
            )}
        </List>
    );
}