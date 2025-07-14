class Resource:
    def __init__(self, id, name, tasks=None):
        self.id = id
        self.name = name
        self.tasks = tasks if tasks is not None else []

    def __repr__(self):
        return f"Resource(id={self.id}, name='{self.name}')"
