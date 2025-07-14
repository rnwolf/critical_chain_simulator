from task import Task
from utils import get_predecessors, get_successors

def find_critical_chain(tasks):
    # Find the longest path in the network
    # This is a simplified approach, a more robust solution would use a proper graph traversal algorithm

    # Find tasks with no predecessors (start tasks)
    start_tasks = [task for task in tasks if not task.predecessors]

    critical_chain = []
    max_duration = 0

    for start_task in start_tasks:
        chain, duration = _find_longest_path(start_task, tasks)
        if duration > max_duration:
            max_duration = duration
            critical_chain = chain

    return critical_chain

def _find_longest_path(task, tasks):
    successors = get_successors(task, tasks)
    if not successors:
        return [task], task.duration

    longest_path = []
    max_duration = 0

    for successor in successors:
        path, duration = _find_longest_path(successor, tasks)
        if duration > max_duration:
            max_duration = duration
            longest_path = path

    return [task] + longest_path, task.duration + max_duration

def schedule_tasks(tasks, critical_chain):
    # Schedule tasks on the critical chain
    current_time = 0
    for task in critical_chain:
        task.start_time = current_time
        task.finish_time = current_time + task.duration
        current_time = task.finish_time

    # Schedule non-critical tasks
    # This is a simplified approach, a more robust solution would handle resource contention
    for task in tasks:
        if task not in critical_chain:
            predecessors = get_predecessors(task, tasks)
            if predecessors:
                task.start_time = max(p.finish_time for p in predecessors)
                task.finish_time = task.start_time + task.duration
            else:
                # For simplicity, schedule tasks with no predecessors at time 0
                task.start_time = 0
                task.finish_time = task.duration

def insert_buffers(tasks, critical_chain):
    # Project buffer
    project_buffer_duration = sum(t.nominal_duration - t.duration for t in critical_chain) / 2
    project_buffer = Task(id=-1, title="Project Buffer", duration=project_buffer_duration, nominal_duration=project_buffer_duration, resources=[], predecessors=[critical_chain[-1].id])
    tasks.append(project_buffer)

    # Feeding buffers
    # This is a simplified approach, a more robust solution would identify all feeding chains
    for task in critical_chain:
        predecessors = get_predecessors(task, tasks)
        for p in predecessors:
            if p not in critical_chain:
                feeding_buffer_duration = sum(t.nominal_duration - t.duration for t in _get_feeding_chain(p, tasks, critical_chain)) / 2
                if feeding_buffer_duration > 0:
                    feeding_buffer = Task(id=-2, title=f"Feeding Buffer for T{task.id}", duration=feeding_buffer_duration, nominal_duration=feeding_buffer_duration, resources=[], predecessors=[p.id])
                    tasks.append(feeding_buffer)
                    task.predecessors.append(feeding_buffer.id)


def _get_feeding_chain(task, tasks, critical_chain):
    # This is a simplified approach to identify a feeding chain
    chain = []
    current_task = task
    while current_task and current_task not in critical_chain:
        chain.insert(0, current_task)
        predecessors = get_predecessors(current_task, tasks)
        if predecessors:
            # For simplicity, we only consider the first predecessor
            current_task = predecessors[0]
        else:
            current_task = None
    return chain
