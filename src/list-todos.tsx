// src/list-todos.tsx
import { ActionPanel, Action, List, showToast, Toast, getPreferenceValues, Icon } from "@raycast/api";
import { useEffect, useState } from "react";
import fetch from "node-fetch";

// Interfaces
interface Preferences {
    token: string;
}

interface TaskList {
    id: string;
    displayName: string;
}

interface Todo {
    id: string;
    title: string;
    status: string;
}

interface GroupedTodos {
    list: TaskList;
    todos: Todo[];
}

// API Call to fetch To-Do lists
async function fetchTaskLists(): Promise<TaskList[]> {
    const { token } = getPreferenceValues<Preferences>();
    try {
        const response = await fetch("https://graph.microsoft.com/v1.0/me/todo/lists", {
            headers: { Authorization: `Bearer ${token}` },
        });
        if (!response.ok) {
            const data = await response.json();
            throw new Error(data.error?.message || "Failed to fetch task lists");
        }
        const data: any = await response.json();
        return data.value.filter((list: any) => list.displayName !== "Flagged Emails");
    } catch (error) {
        throw new Error(error instanceof Error ? error.message : "Could not fetch task lists");
    }
}

// API Call to fetch To-Dos for a specific list
async function fetchTodosForList(listId: string): Promise<Todo[]> {
    const { token } = getPreferenceValues<Preferences>();
    try {
        const response = await fetch(`https://graph.microsoft.com/v1.0/me/todo/lists/${listId}/tasks?$filter=status ne 'completed'`, {
            headers: { Authorization: `Bearer ${token}` },
        });
        if (!response.ok) {
            const data = await response.json();
            // We log the error here but return an empty array to not block the entire process if one list fails
            console.error(`Throttling or error on list ${listId}: ${data.error?.message}`);
            return [];
        }
        const data: any = await response.json();
        return data.value;
    } catch (error) {
        console.error(`Failed to fetch todos for list ${listId}:`, error);
        return [];
    }
}

// API call to mark a task as complete
async function markTaskAsCompleteAPI(listId: string, taskId: string) {
    const { token } = getPreferenceValues<Preferences>();
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

    // THIS IS THE CORRECTED FUNCTION
    async function loadTodos() {
        setIsLoading(true);
        const toast = await showToast({ style: Toast.Style.Animated, title: "Loading tasks..." });
        try {
            const taskLists = await fetchTaskLists();
            const allGroupedTodos: GroupedTodos[] = [];

            // We use a sequential for...of loop here to prevent sending too many requests at once.
            // This is the fix for the "throttled" error.
            for (const list of taskLists) {
                const todos = await fetchTodosForList(list.id);
                if (todos.length > 0) {
                    allGroupedTodos.push({ list, todos });
                }
            }

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

        const newGroupedTodos = groupedTodos.map(group => {
            if (group.list.id === listId) {
                return { ...group, todos: group.todos.filter(t => t.id !== taskId) };
            }
            return group;
        }).filter(group => group.todos.length > 0);
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