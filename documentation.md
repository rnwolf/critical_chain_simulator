# CCPM Scheduler Algorithm Documentation

This document provides a detailed explanation of the CCPM (Critical Chain Project Management) scheduling algorithm implemented in the VBA codebase.

## Table of Contents

* [Data Structures](#data-structures)
  * [Tache (Task) Class](#tache-task-class)
  * [Ressource (Resource) Class](#ressource-resource-class)
* [Scheduling Algorithm](#scheduling-algorithm)
  * [1. Initialization](#1-initialization)
  * [2. Critical Chain Identification](#2-critical-chain-identification)
  * [3. Positioning of Tasks Linked to the Critical Chain](#3-positioning-of-tasks-linked-to-the-critical-chain)
  * [4. Positioning of the Last Unsorted Tasks](#4-positioning-of-the-last-unsorted-tasks)
  * [5. Buffer Generation](#5-buffer-generation)
  * [6. GANTT Chart Display](#6-gantt-chart-display)
* [Buffer Consumption](#buffer-consumption)
* [Utility Functions](#utility-functions)

## Data Structures

The algorithm uses two main data structures to represent tasks and resources: `Tache` and `Ressource`.

### Tache (Task) Class

The `Tache` class (`VBA_export/Class Modules/Tache.cls`) represents a single task in the project.

**Properties:**

*   `id` (Integer): A unique identifier for the task.
*   `intitule` (String): The title or description of the task.
*   `duree` (Integer): The estimated duration of the task.
*   `duree_nominale` (Integer): The nominal duration of the task (used for buffer calculation).
*   `ress` (String): A string containing the IDs of the resources required for the task.
*   `preds` (String): A comma-separated string of predecessor task IDs.
*   `fin` (String): The scheduled finish time of the task.
*   `debut` (String): The scheduled start time of the task.
*   `t` (Integer): The type of the task. The possible values are:
    *   `1`: Critical chain task
    *   `2`: Secondary chain task
    *   `3`: Free task
    *   `4`: Buffer task
*   `debut_reel` (Integer): The actual start time of the task (used for progress tracking).
*   `colonne_fin_reelle` (Integer): The column number in the Excel sheet that represents the actual finish time.

### Ressource (Resource) Class

The `Ressource` class (`VBA_export/Class Modules/Ressource.cls`) represents a resource that can be assigned to tasks.

**Properties:**

*   `id` (Integer): A unique identifier for the resource.
*   `nom` (String): The name of the resource.
*   `t` (String): A string containing the IDs of the tasks assigned to the resource.

## Scheduling Algorithm

The main scheduling algorithm is implemented in the `ordo` subroutine in the `ordonnancement` module (`VBA_export/Modules/ordonnancement.bas`).

The algorithm follows these steps:

### 1. Initialization

The `ordo` subroutine starts by initializing some variables and clearing the "LOGS" worksheet. It then calls the `retrieve_tasks` subroutine to load the tasks from the "TACHES" worksheet into a collection of `Tache` objects.

```vba
Sub ordo(Optional d As Integer = 0)

    planning_alert = False

    If d <> 1 Then
        Dim p As Integer, sh As Worksheet
        Set sh = ThisWorkbook.Worksheets("LOGS")
        While sh.Cells(22 + p, 9).value <> ""
            p = p + 1
        Wend
        sh.Range("O15:P200").Clear
        sh.Range(sh.Cells(22, 9), sh.Cells(24 + p, 11)).Clear
        Dim re_plan As Boolean: re_plan = False
        Call retrieve_tasks
    End If
    Dim t_triees As Collection, s As Worksheet, ta As Integer
    Set s = ThisWorkbook.Worksheets("GANTT")
    Set t_triees = New Collection 'Sorted task tables
```

### 2. Critical Chain Identification

The algorithm identifies the critical chain by finding the longest path of dependent tasks.

It starts by finding a "first task" (a task with no predecessors). It then iterates through the tasks, calculating the longest path to find the critical chain. The tasks in the critical chain are stored in the `chaine_critique` (critical_chain) collection.

```vba
    '--------- Critical chain generation ---------'

    'What task should I start with? Necessarily a "first task" - > any predecessor
    Dim j As Integer, indice_max As Integer, max As Integer, previous_max As Integer, ante As Object
    previous_max = 0
    indice_max = 1
    For j = 1 To taches.Count
        If taches(j).get_preds = "" Then
            max = taches(j).get_duree
            Dim z As Tache
            Set z = taches(j)
            While antecedants(z, taches).Count >= 1 'As long as there are still some predecessors

                Set ante = antecedants(z, taches) 'Recover the antecdants of the last sorted task
                max = max + ante(next_tache(ante)).get_duree()
                If max > 10000 Then
                    MsgBox " Please check the previous entries."
                    Exit Sub
                End If
                Set z = ante(next_tache(ante))

            Wend
            If previous_max > 0 Then
                If previous_max < max Then
                    indice_max = j
                    previous_max = max
                End If
            Else
                indice_max = j
                previous_max = max
            End If
        End If
    Next j

    t_triees.Add taches(indice_max)
    taches.Remove indice_max

    t_triees(1).set_debut (0)
    t_triees(1).set_fin (t_triees(1).get_duree)
    t_triees(1).set_type (1)

    While antecedants(t_triees(t_triees.Count), taches).Count >= 1 'As long as there are still tasks to be added in a critical chain
        'Dim ante As Object
        Set ante = antecedants(t_triees(t_triees.Count), taches) 'Recover the predecessors of the last sorted task
        t_triees.Add ante(next_tache(ante)) 'the one with the longest duration is recovered from among these predecessors
        Call remove_task_by_id(t_triees(t_triees.Count).get_ID, taches) 'the task is removed from the unsorted tasks

        t_triees(t_triees.Count).set_debut (t_triees(t_triees.Count - 1).get_fin)
        t_triees(t_triees.Count).set_fin (CInt(t_triees(t_triees.Count).get_debut) + CInt(t_triees(t_triees.Count).get_duree))
        t_triees(t_triees.Count).set_type (1) 'critical chain type

    Wend
```

### 3. Positioning of Tasks Linked to the Critical Chain

The algorithm then positions the tasks that are not on the critical chain but are linked to it. It uses a recursive function called `recursive_positioning` to place these tasks.

The `recursive_positioning` function takes the following arguments:

*   `k` (Integer): The index of the current task in the `t_triees` collection.
*   `t` (Collection): The collection of sorted tasks.
*   `s_chain` (Collection): A collection to store the tasks in the secondary chain.
*   `critical_chain` (Collection): The collection of tasks in the critical chain.
*   `alerte_on` (Integer): An optional argument to enable or disable planning alerts.

The function works by recursively traversing the predecessors of the current task and positioning them on the schedule. It uses the `set_intermediate_task` function to place each task, taking into account resource constraints.

```vba
'--------- Positioning of tasks linked to the critical chain ---------'

    'The remaining tasks in "tasks" are not yet sorted
    'they are positioned around the critical chain while avoiding overlapping resources

Dim i As Integer, k As Integer, end_loop As Integer, secondary_chains As Collection 'secondary_chains is meant to collect all 2ndary chains
end_loop = t_triees.Count
Set secondary_chains = New Collection

Dim previous_size As Integer
previous_size = chaine_critique.Count

For i = 2 To end_loop
    k = end_loop + 2 - i 'we travel in the opposite direction to avoid being affected by the insertion

    Dim s_chain As Collection
    Set s_chain = New Collection 'one secondary chain

    If d = 1 Then
        Call recursive_positioning(k, t_triees, s_chain, chaine_critique, 1) 'launch recursion with alert monitoring
    Else
        Call recursive_positioning(k, t_triees, s_chain, chaine_critique)
    End If

    'saving secondary chain only if it has tasks
    If s_chain.Count > 0 Then
        secondary_chains.Add s_chain
    End If
Next i
```

### 4. Positioning of the Last Unsorted Tasks

After positioning the tasks linked to the critical chain, the algorithm positions any remaining unsorted tasks. These are tasks that have no predecessors or whose predecessors have not yet been scheduled.

The algorithm iterates through the remaining tasks and places them on the schedule using the `set_intermediate_task` function.

```vba
'--------- Positioning of the last unsorted tasks ---------'

    'At this stage, there may still be unsorted tasks (present in the "tasks" collection)
    'they potentially have one or more predecessors : we know their left limit
    'If they have no predecessors, they can be placed in sync with the last task (Non-overlapping resource condition)
'If d = 1 Then 'We do it only if buffers already generated
    Dim free_chains As Collection 'storage of the "free" chains that will be generated
    Dim counter As Integer: counter = 0
    Set free_chains = New Collection

    'as long as there are still unsorted tasks
    While taches.Count > 0
        Dim Target As Integer
        Target = 1
        For i = 2 To taches.Count
            'trying to focus first on the tasks that have no predecessors
            If antecedants(taches(i), taches).Count = 0 Then
               Target = i
            End If
        Next i
            'We have selected a task without a predecessor, we will position it

        Call set_intermediate_task(t_triees, t_triees.Count - counter, taches(Target), last_task_indice(t_triees), max_preds_end(taches(Target), t_triees))
        counter = counter + 1
        taches(Target).set_type (3) 'free type
        t_triees.Add taches(Target)
        taches.Remove Target

            'We have placed a starting task for potential recursions

        'run recursion if preceded
        If t_triees(t_triees.Count).get_preds <> "" Then
            Dim f_chain As Collection
            Set f_chain = New Collection
            Call recursive_positioning(t_triees.Count, t_triees, f_chain, chaine_critique)

            If f_chain.Count > 0 Then
                free_chains.Add f_chain
            End If
        End If

    Wend
```

### 5. Buffer Generation

The `generate_buffers` subroutine in the `buffers` module (`VBA_export/Modules/buffers.bas`) calculates and creates buffers for the critical chain and secondary chains.

*   **Critical Chain Buffer**: The buffer for the critical chain is calculated by summing the differences between the nominal duration and the actual duration of each task in the chain. A new task is then created to represent the critical chain buffer.
*   **Secondary Chain Buffers**: The subroutine iterates through the secondary chains, calculates their buffers, and creates buffer tasks for them.
*   **Link Breaking and Remaking**: The predecessors of the tasks are modified to insert the buffers into the schedule. The link between the critical chain and the first task of a secondary chain is broken, and the buffer task is inserted in between.

```vba
'calculation of buffer duration for each chain, preparation for positioning in the GANTT
'"cc" : critical chain, "sc" : collection of secondary chains
Sub generate_buffers(cc As Collection, sc As Collection, Optional alert As Integer = 0)

    Dim i As Integer, j As Integer, t As Tache, sum As Integer
    Dim indice_max As Integer: indice_max = cc.Count
    For i = 1 To cc.Count
        sum = sum + (cc(i).get_duree_nominale - cc(i).get_duree) 'Sum of nominal durations - optimized
        If cc(i).get_fin > cc(indice_max).get_fin Then
            indice_max = i
        End If
    Next i

    'creation of a task that plays the role of the critical chain buffer
    Set t = New Tache
    t.set_attributes "Critical chain buffer", sum, "1", CStr(cc(indice_max).get_ID)
    t.set_type 4
    taches.Add t
    ThisWorkbook.Worksheets("LOGS").Cells(15, 16).value = sum 'the length of the buffer is recorded

    ' ... (code for secondary chain buffers) ...
End Sub
```

### 6. GANTT Chart Display

Finally, the `affichage_GANTT` (display_GANTT) subroutine (in the `affichages` module) is called to display the schedule on a GANTT chart. This subroutine is not documented here, but it is responsible for rendering the schedule in the "GANTT" worksheet.

## Buffer Consumption

The `consume_buffers` subroutine in the `buffers` module (`VBA_export/Modules/buffers.bas`) calculates the percentage of the buffer that has been consumed and updates the "LOGS_FV_CHART" worksheet to display the progress on a fever chart.

*   It retrieves the chains of tasks using `retrieve_chains`.
*   It iterates through each chain and calculates the total duration and the completed duration.
*   It calculates the buffer consumption based on the difference between the planned progress and the actual progress.
*   It updates the "LOGS_FV_CHART" worksheet with the data for the fever chart.

## Utility Functions

The `utils` module (`VBA_export/Modules/utils.bas`) contains several utility functions that are used throughout the application.

*   `retrieve_tasks`: Reads task data from the "TACHES" worksheet and creates a collection of `Tache` objects.
*   `retrieve_ressources`: Reads resource data from the "TACHES" worksheet and creates a collection of `Ressource` objects.
*   `retrieve_chains`: Retrieves the chains of tasks from the "LOGS" worksheet.
*   `antecedants`: Returns a collection of tasks that are the antecedents (predecessors) of a given task.
*   `last_task` and `last_task_indice`: Find the last task in a collection of tasks.
*   `task_in_tab_by_id` and `get_task_index_by_id`: Find a task in a collection by its ID.
*   `dans_chaine_critique` and `dans_quel_chaine`: Check if a task is in the critical chain or in which chain it is.
