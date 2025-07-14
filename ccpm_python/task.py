class Task:
    def __init__(self, id, title, duration, nominal_duration, resources, predecessors, task_type=0, actual_start=0, actual_finish_column=0):
        self.id = id
        self.title = title
        self.duration = duration
        self.nominal_duration = nominal_duration
        self.resources = resources
        self.predecessors = predecessors
        self.finish_time = 0
        self.start_time = 0
        self.task_type = task_type
        self.actual_start = actual_start
        self.actual_finish_column = actual_finish_column

    def __repr__(self):
        return f"Task(id={self.id}, title='{self.title}')"
