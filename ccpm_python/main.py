from utils import retrieve_tasks, retrieve_resources
from ccpm import find_critical_chain, schedule_tasks, insert_buffers

def main():
    # Retrieve tasks and resources from the input file
    tasks = retrieve_tasks()
    resources = retrieve_resources()

    # Find the critical chain
    critical_chain = find_critical_chain(tasks)

    # Insert buffers
    insert_buffers(tasks, critical_chain)

    # Schedule all tasks
    schedule_tasks(tasks, critical_chain)

    # Write the output to a file
    with open("output.txt", "w") as f:
        f.write("# Scheduled Tasks\n")
        for task in sorted(tasks, key=lambda t: t.start_time):
            f.write(f"Task {task.id}: {task.title}, Start: {task.start_time}, Finish: {task.finish_time}\n")

        f.write("\n# Critical Chain\n")
        for task in critical_chain:
            f.write(f"Task {task.id}: {task.title}\n")

if __name__ == "__main__":
    main()
