import sys


def do_work(task_id):
    print(f'o{task_id} complete')


if __name__ == '__main__':
    if len(sys.argv) != 2:
        print('Please provide one argument', file=sys.stderr)
        exit(1)
    try:
        task_id = sys.argv[1]  # Get the task id from the command line argument
        do_work(task_id)
    except Exception as e:
        print(e)
