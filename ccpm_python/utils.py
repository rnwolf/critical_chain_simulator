from task import Task
from resource import Resource

def retrieve_tasks(filename="input.txt"):
    tasks = []
    with open(filename, 'r') as f:
        lines = f.readlines()
        task_section = False
        for line in lines:
            if line.strip() == "# Tasks":
                task_section = True
                continue
            elif line.strip() == "# Resources":
                task_section = False
                break

            if task_section and line.strip() and not line.startswith('#'):
                parts = line.strip().split(',')
                id = int(parts[0])
                title = parts[1]
                duration = int(parts[2])
                nominal_duration = int(parts[3])
                resources = [int(r) for r in parts[4].split(';') if r]
                predecessors = [int(p) for p in parts[5].split(';') if p] if len(parts) > 5 else []
                tasks.append(Task(id, title, duration, nominal_duration, resources, predecessors))
    return tasks

def retrieve_resources(filename="input.txt"):
    resources = []
    with open(filename, 'r') as f:
        lines = f.readlines()
        resource_section = False
        for line in lines:
            if line.strip() == "# Resources":
                resource_section = True
                continue

            if resource_section and line.strip() and not line.startswith('#'):
                parts = line.strip().split(',')
                id = int(parts[0])
                name = parts[1]
                resources.append(Resource(id, name))
    return resources

def get_task_by_id(tasks, task_id):
    for task in tasks:
        if task.id == task_id:
            return task
    return None

def get_predecessors(task, tasks):
    predecessors = []
    for pred_id in task.predecessors:
        predecessor = get_task_by_id(tasks, pred_id)
        if predecessor:
            predecessors.append(predecessor)
    return predecessors

def get_successors(task, tasks):
    successors = []
    for t in tasks:
        if task.id in t.predecessors:
            successors.append(t)
    return successors
