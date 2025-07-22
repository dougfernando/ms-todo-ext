// src/create-todo.tsx
import {
    ActionPanel,
    Form,
    Action,
    showToast,
    Toast,
    closeMainWindow,
    useNavigation,
} from "@raycast/api";
import { useEffect, useState } from "react";
import fetch from "node-fetch";
import { authorize, getAccessToken } from "./auth";

// Interfaces
interface TaskList {
    id: string;
    displayName: string;
}

interface TaskForm {
    title: string;
    content: string;
    dueDateTime: Date | null;
    isComplete: boolean;
    importance: string;
    taskList: string;
}

interface TaskResponse {
    id: string;
}

// API Requests
async function fetchTaskLists(): Promise<TaskList[]> {
    const token = await getAccessToken();
    try {
        const response = await fetch("https://graph.microsoft.com/v1.0/me/todo/lists", {
            headers: { Authorization: `Bearer ${token}` },
        });
        const data: any = await response.json();

        if (data.error) {
            showToast(Toast.Style.Failure, "Error fetching task lists", data.error.message);
            return [];
        }
        return data.value.filter((taskList: TaskList) => taskList.displayName !== "Flagged Emails");
    } catch (error) {
        const message = error instanceof Error ? error.message : "An unknown error occurred";
        showToast(Toast.Style.Failure, "Error", message);
        console.log(error);
        return [];
    }
}

async function createTodo(task: TaskForm): Promise<TaskResponse> {
    const token = await getAccessToken();
    const body: any = {
        title: task.title,
        body: { content: task.content, contentType: "text" },
        importance: task.importance,
    };

    if (task.dueDateTime) {
        body.dueDateTime = {
            dateTime: task.dueDateTime.toISOString(),
            timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone,
        };
    }

    if (task.isComplete) {
        body.status = "completed";
    }

    const response = await fetch(`https://graph.microsoft.com/v1.0/me/todo/lists/${task.taskList}/tasks`, {
        method: "POST",
        headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
        body: JSON.stringify(body),
    });

    if (!response.ok) {
        const errorData: any = await response.json();
        console.log(errorData);
        throw new Error(errorData.error?.message || "Failed to create task");
    }

    return (await response.json()) as TaskResponse;
}

export default function CreateTodoCommand() {
    const [taskLists, setTaskLists] = useState<TaskList[]>([]);
    const [isLoading, setIsLoading] = useState(true);

    useEffect(() => {
        async function fetchData() {
            await authorize();
            setIsLoading(true);
            const taskListsResponse = await fetchTaskLists();
            setTaskLists(taskListsResponse);
            setIsLoading(false);
        }
        fetchData();
    }, []);

    // UPDATED: Now pops to root before closing the window.
    async function handleSubmit(values: TaskForm) {
        const toast = await showToast({ style: Toast.Style.Animated, title: "Creating task..." });
        try {
            await createTodo(values);

            toast.style = Toast.Style.Success;
            toast.title = "Task Created";

            // Then, close the main window.
            await closeMainWindow({ clearRootSearch: true });

        } catch (error) {
            toast.style = Toast.Style.Failure;
            toast.title = "Error";
            toast.message = error instanceof Error ? error.message : "Could not create task";
            console.log(error);
        }
    }

    return (
        <Form
            isLoading={isLoading}
            actions={
                <ActionPanel>
                    <Action.SubmitForm title="Create To-Do" onSubmit={handleSubmit} />
                </ActionPanel>
            }
        >
            <Form.TextField id="title" title="Title" placeholder="Buy milk" />
            <Form.TextArea id="content" title="Notes" placeholder="More details about the task" />
            <Form.Separator />
            <Form.Dropdown id="taskList" title="Task List">
                {taskLists.map((list) => (
                    <Form.Dropdown.Item key={list.id} value={list.id} title={list.displayName} />
                ))}
            </Form.Dropdown>
            <Form.Separator />
            <Form.DatePicker id="dueDateTime" title="Due Date" type={Form.DatePicker.Type.Date} />
            <Form.Dropdown id="importance" title="Importance" defaultValue="normal">
                <Form.Dropdown.Item value={"low"} title={"Low"} />
                <Form.Dropdown.Item value={"normal"} title={"Normal"} />
                <Form.Dropdown.Item value={"high"} title={"High"} />
            </Form.Dropdown>
            <Form.Checkbox id="isComplete" label="Mark this task complete" />
        </Form>
    );
}