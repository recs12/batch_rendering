"""
Only works in mode TC.
"""

import datetime
import getpass
import os
import pathlib
import subprocess
import time

import click
import pandas as pd
from loguru import logger

# import notifiers

# params = {
#     "username": "slimane.rechdi@gmail.com",
#     "password": "abc123", # prompt with click
#     "to": "dest@gmail.com"
# }

# Send a single notification
# notifier = notifiers.get_notifier("gmail")
# notifier.notify(message="The application is running!", **params)

user_name = getpass.getuser()

def render_one_item(items, user, password, folder_target, group, role, server, mode):
    print(f"[INFO] TC - Part Number: {items[1]}")
    item_id, revision = items[0], items[1]
    print(f"item {item_id} / {revision}")
    p1 = subprocess.run(
        f"SEToolRender {mode} -i {item_id} -v {revision} -u {user} -p {password} -g {group} -r {role} -s {server} -o {folder_target}",
        shell=True,
        check=False,
        capture_output=True,
        text=True,  # display the main macro output
        stderr=subprocess.DEVNULL,  # ignore the errors
    )

    # Logging of the error message.
    if p1.returncode != 0:
        logger.error(f"[ERROR] : {p1.stderr}")

    # check the presence of the files (json, bmp)
    bmp = os.path.join(folder_target, f"{item_id}-Rev-{revision}.bmp")
    json = os.path.join(folder_target, f"{item_id}-Rev-{revision}.json")
    file_bmp = pathlib.Path(bmp)
    file_json = pathlib.Path(json)
    if file_bmp.exists() and file_json.exists():
        logger.info(f"[OK   ] : {item_id} Files downloaded")
    else:
        logger.error(f"[FAIL ] : {item_id} Files not downloaded")

    # If server disconnection code 666 then tempo of 15 min.
    if p1 == 666:
        time.sleep(60*15)


# maybe we could add some defaults values.
@click.command(help="Ctrl + C to stop the macro.")
@click.argument("excel", nargs=1, required=True, type=click.Path(exists=True))
@click.option("--user", "-u", nargs=1, required=False, type=str, help="Your acronym.")
@click.password_option()
@click.option(
    "--folder_target",
    "-o",
    nargs=1,
    required=True,
    help="Location where to download the pictures.",
    type=click.Path(exists=True),
)
@click.option(
    "--group",
    "-g",
    default="Engineering",
    show_default=True,
    nargs=1,
    required=False,
    type=click.Choice(["Engineering"]),
    help="Group optional.",
)
@click.option(
    "--role",
    "-r",
    default="Designer",
    show_default=True,
    nargs=1,
    required=False,
    type=click.Choice(["Designer"]),
    help="Role optional.",
)
@click.option(
    "--server",
    "-s",
    default="TC_PROD",
    show_default=True,
    nargs=1,
    required=False,
    type=click.Choice(["TC_PROD"]),
    help="Server optional.",
)
@click.option(
    "--mode",
    "-m",
    default="TC",
    show_default=True,
    nargs=1,
    required=False,
    type=click.Choice(["TC"]),
    help="TC is the only mode enabled.",
)
@click.option(
    "--debug-mode",
    default="ERROR",
    required=False,
    type=click.Choice(["ERROR", "INFO"]),
    help="Debug mode for logging file.",
)
def batch_rendering(
    excel,
    password,
    folder_target,
    group,
    role,
    server,
    mode,
    debug_mode,
    user=user_name,
):
    """Run SEToolRendering in batch with an excel file."""

    logger.add(
        "rendering_error.log",
        format="{time:YYYY-MM-DD at HH:mm:ss} | {level} | {message}",
        level=debug_mode,
    )
    print(f"Logging mode: {debug_mode}")
    df = pd.read_excel(
        excel,
        index_col=None,
        dtype={
            "ID": str,
            "Revision": str,
        },  # names of columns in excel must be excatly the same.
    )
    # Define the shape of the excel file.
    number_rows, number_columns = df.shape

    try:
        tic = time.perf_counter()
        # For each row run a command line.
        for row in range(number_rows):
            items: list = [df.iat[row, column] for column in range(number_columns)]
            render_one_item(
                items, user, password, folder_target, group, role, server, mode
            )
            logger.info(f"[ITERATION] {row+1}/{number_rows}")

    except subprocess.CalledProcessError:
        logger.exception({row + 1})

    finally:
        toc = time.perf_counter()
        seconds = toc - tic
        logger.info(f"Time to completion: {datetime.timedelta(seconds=seconds)}")
