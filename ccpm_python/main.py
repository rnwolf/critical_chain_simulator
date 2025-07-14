import os
import csv
from utils import retrieve_tasks, retrieve_resources
from ccpm import find_critical_chain, schedule_tasks, insert_buffers

def main():
    # The script is run from the root directory, so the input file is in the ccpm_python directory
    input_file = "ccpm_python/input.txt"
    output_file = "ccpm_python/output.csv"

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
    with open(output_file, "w", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["ID", "Title", "Start Time", "Finish Time", "Resources", "Predecessors", "Critical Chain"])
        for task in sorted(tasks, key=lambda t: t.start_time):
            writer.writerow([
                task.id,
                task.title,
                task.start_time,
                task.finish_time,
                ";".join(map(str, task.resources)),
                ";".join(map(str, task.predecessors)),
                task in critical_chain
            ])

if __name__ == "__main__":
    main()
