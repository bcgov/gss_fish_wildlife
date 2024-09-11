import subprocess
import multiprocessing as mp
import sys

NUMBER_OF_JOBS= 4  # Define the number of tasks to run in parallel


def batch_ast():
    print(f'Batch Ast Running')

def run_worker():
    # Simulate the worker's job directly

    print("Running a job")
    batch_ast()





if __name__ == '__main__':

    # Run the multiprocessing tasks if no command line arguments are provided
    pool = mp.Pool(NUMBER_OF_JOBS)  # Create a pool of workers

    # Schedule tasks asynchronously
    for task_number in range(NUMBER_OF_JOBS):
        # pool.apply_async(run_worker, (str(task_number + 1),))  # Pass task numbers as arguments
        pool.apply_async(
        run_worker,          # The function to run
        args=(),       # Positional arguments for run worker
        #kwed={}, 
        # callback=handle_result,  # Callback function for handling results
        # error_callback=handle_error  # Callback function for handling errors
    )

    pool.close()  # Close the pool to prevent any more tasks from being submitted
    pool.join()   # Wait for all tasks to complete
    

