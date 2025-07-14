import os
from utils import retrieve_tasks, retrieve_resources
from ccpm import find_critical_chain, schedule_tasks, insert_buffers

def main():
    # The script is run from the root directory, so the input file is in the ccpm_python directory
    input_file = "ccpm_python/input.txt"
    output_file = "ccpm_python/output.txt"

    # Retrieve tasks and resources from the input file
    tasks = retrieve_tasks("ccpm_python")
    resources = retrieve_resources("ccpm_python")

    # Find the critical chain
    critical_chain = find_critical_chain(tasks)

    # Insert buffers
    insert_buffers(tasks, critical_chain)

    # Schedule all tasks
    schedule_tasks(tasks, critical_chain)

    # Write the output to a file
    with open(output_file, "w") as f:
        f.write("# Scheduled Tasks\n")
        for task in sorted(tasks, key=lambda t: t.start_time):
            f.write(f"Task {task.id}: {task.title}, Start: {task.start_time}, Finish: {task.finish_time}\n")

        f.write("\n# Critical Chain\n")
        for task in critical_chain:
            f.write(f"Task {task.id}: {task.title}\n")

if __name__ == "__main__":
    main()
