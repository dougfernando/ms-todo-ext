import { ActionPanel, Action, List, showToast, Toast, Icon } from "@raycast/api";
import { useEffect, useState } from "react";
import fetch from "node-fetch";
import { authorize, getAccessToken } from "./auth";

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

type FilterType = 'all' | 'with-due-date' | 'important';

// Filter function to apply filters to todos
function applyFilters(todos: Todo[], filterType: FilterType): Todo[] {
  switch (filterType) {
    case 'with-due-date':
      return todos.filter(todo => todo.dueDateTime);
    case 'important':
      return todos.filter(todo => todo.importance === 'high');
    case 'all':
    default:
      return todos;
  }
}

// Interfaces
interface TaskList {
  id: string;
  displayName: string;
}

interface Todo {
  id: string;
  title: string;
  status: string;
  importance: 'low' | 'normal' | 'high';
  dueDateTime?: {
    dateTime: string;
    timeZone: string;
  };
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

// API call to fetch To-Dos for a specific list
async function fetchTodosForList(listId: string): Promise<Todo[]> {
  const token = await getAccessToken();
  try {
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/me/todo/lists/${listId}/tasks?$filter=status ne 'completed'`,
      {
        headers: { Authorization: `Bearer ${token}` },
      }
    );
    if (!response.ok) {
      const data = (await response.json()) as { error?: { message: string } };
      throw new Error(data.error?.message || `HTTP error ${response.status}`);
    }
    const data: any = await response.json();
    return data.value;
  } catch (error) {
    throw new Error(error instanceof Error ? error.message : "Could not fetch tasks for list");
  }
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

export default function ListTasksByListCommand() {
  const [taskLists, setTaskLists] = useState<TaskList[]>([]);
  const [selectedList, setSelectedList] = useState<TaskList | null>(null);
  const [todos, setTodos] = useState<Todo[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [currentFilter, setCurrentFilter] = useState<FilterType>('all');

  async function loadTaskLists() {
    await authorize();
    setIsLoading(true);
    const toast = await showToast({ style: Toast.Style.Animated, title: "Loading lists..." });
    try {
      const lists = await fetchTaskLists();
      setTaskLists(lists);
      toast.style = Toast.Style.Success;
      toast.title = "Lists Loaded";
    } catch (error) {
      toast.style = Toast.Style.Failure;
      toast.title = "Error";
      toast.message = error instanceof Error ? error.message : "Could not load lists";
    } finally {
      setIsLoading(false);
    }
  }

  async function loadTodosForList(list: TaskList) {
    setSelectedList(list);
    setIsLoading(true);
    const toast = await showToast({ style: Toast.Style.Animated, title: `Loading tasks for ${list.displayName}...` });
    try {
      const fetchedTodos = await fetchTodosForList(list.id);
      setTodos(fetchedTodos);
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
    loadTaskLists();
  }, []);

  async function handleMarkAsComplete(taskId: string) {
    if (!selectedList) return;

    const originalTodos = [...todos];
    const newTodos = todos.filter((t) => t.id !== taskId);
    setTodos(newTodos);

    try {
      const response = await markTaskAsCompleteAPI(selectedList.id, taskId);
      if (!response.ok) {
        setTodos(originalTodos);
        const errorData: any = await response.json();
        await showToast(Toast.Style.Failure, "Failed to Complete Task", errorData.error?.message);
      } else {
        await showToast(Toast.Style.Success, "Task Completed!");
      }
    } catch (error) {
      setTodos(originalTodos);
      const message = error instanceof Error ? error.message : "An unknown error occurred";
      await showToast(Toast.Style.Failure, "Error", message);
    }
  }

  function handleBack() {
    setSelectedList(null);
    setTodos([]);
  }

  if (selectedList) {
    const filteredTodos = applyFilters(todos, currentFilter);
    
    return (
      <List 
        isLoading={isLoading} 
        navigationTitle={selectedList.displayName} 
        searchBarPlaceholder="Filter your to-dos..."
        searchBarAccessory={
          <List.Dropdown
            tooltip="Filter Tasks"
            value={currentFilter}
            onChange={(newFilter) => setCurrentFilter(newFilter as FilterType)}
          >
            <List.Dropdown.Item title="All Tasks" value="all" icon={Icon.List} />
            <List.Dropdown.Item title="With Due Date" value="with-due-date" icon={Icon.Calendar} />
            <List.Dropdown.Item title="Important" value="important" icon={Icon.Important} />
          </List.Dropdown>
        }
      >
        {filteredTodos.length === 0 && !isLoading ? (
          <List.EmptyView 
            title={currentFilter === 'all' ? "No To-Dos Found" : "No Matching Tasks"} 
            description={currentFilter === 'all' ? "You're all caught up!" : `No tasks match the "${currentFilter === 'with-due-date' ? 'With Due Date' : 'Important'}" filter`} 
            icon={Icon.Checkmark} 
          />
        ) : (
          <List.Section title="Tasks">
            {filteredTodos.map((todo) => (
              <List.Item
                key={todo.id}
                title={todo.title}
                icon={Icon.Circle}
                accessories={formatDueDate(todo.dueDateTime) ? [{ text: formatDueDate(todo.dueDateTime) }] : undefined}
                actions={
                  <ActionPanel>
                    <Action title="Mark as Complete" icon={Icon.CheckCircle} onAction={() => handleMarkAsComplete(todo.id)} />
                    <Action.OpenInBrowser title="Open in To Do" url="https://to-do.live.com" />
                    <Action title="Back to Lists" icon={Icon.ArrowLeft} onAction={handleBack} />
                    <Action
                      title="Reload"
                      icon={Icon.Repeat}
                      onAction={() => loadTodosForList(selectedList)}
                      shortcut={{ modifiers: ["cmd"], key: "r" }}
                    />
                  </ActionPanel>
                }
              />
            ))}
          </List.Section>
        )}
      </List>
    );
  }

  return (
    <List isLoading={isLoading} searchBarPlaceholder="Filter your lists...">
      {taskLists.length === 0 && !isLoading ? (
        <List.EmptyView title="No Lists Found" icon={Icon.List} />
      ) : (
        <List.Section title="Task Lists">
          {taskLists.map((list) => (
            <List.Item
              key={list.id}
              title={list.displayName}
              icon={Icon.List}
              actions={
                <ActionPanel>
                  <Action title="View Tasks" icon={Icon.ArrowRight} onAction={() => loadTodosForList(list)} />
                  <Action.OpenInBrowser title="Open in To Do" url="https://to-do.live.com" />
                  <Action
                    title="Reload"
                    icon={Icon.Repeat}
                    onAction={loadTaskLists}
                    shortcut={{ modifiers: ["cmd"], key: "r" }}
                  />
                </ActionPanel>
              }
            />
          ))}
        </List.Section>
      )}
    </List>
  );
}