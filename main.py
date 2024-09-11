import subprocess
import multiprocessing as mp

NUMBER_OF_TASKS = 4  # Define the number of tasks to run in parallel


def work(task_id):
    # Run the worker script with the task_id as an argument
    command = ['python', 'worker.py', task_id]
    subprocess.call(command)


if __name__ == '__main__':
    pool = mp.Pool(NUMBER_OF_TASKS)  # Create a pool of workers

    # Schedule tasks asynchronously
    for task_number in range(NUMBER_OF_TASKS):
        pool.apply_async(work, (str(task_number + 1),))  # Pass task numbers as arguments

    pool.close()  # Close the pool to prevent any more tasks from being submitted
    pool.join()   # Wait for all tasks to complete
