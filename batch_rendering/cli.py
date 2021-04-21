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
import psutil

from batch_rendering.ranges import *

import notifiers

params = {
    "username": "slimane.rechdi@gmail.com",
    "password": "GGCaN6$eI2021", # prompt with click
    "to": "slimane.rechdi@gmail.com"
}

user_name = getpass.getuser()

def se_opened_instance():
    """Check the number of instance of solidedge."""
    return [p.name() for p in psutil.process_iter()].count("Edge.exe")

def render_one_item(items, user, password, folder_target, group, role, server, mode):
    print(f"[INFO] TC - Part Number: {items[1]}")
    item_id, revision = items[0], items[1]
    print(f"item {item_id} / {revision}")
    p1 = subprocess.run(
        f"SEToolRender {mode} -i {item_id} -v {revision} -u {user} -p {password} -g {group} -r {role} -s {server} -o {folder_target}",
        shell=True,
        check=False,
        capture_output=False,
        text=True,  # display the main macro output
        stderr=subprocess.DEVNULL,  # ignore the errors
    )

    # Logging of the error message.
    if p1.returncode != 0:
        logger.error(f"[ERROR] : {p1.stdout}")

    # check the presence of the files (json, bmp)
    bmp = os.path.join(folder_target, f"{item_id}-Rev-{revision}.bmp")
    json = os.path.join(folder_target, f"{item_id}-Rev-{revision}.json")
    file_bmp = pathlib.Path(bmp)
    file_json = pathlib.Path(json)

    if file_bmp.exists():
        logger.info(f"[OK   ] : {item_id} .bmp downloaded")
    else:
        logger.error(f"[FAIL ] : {item_id} .bmp not downloaded")

    if file_json.exists():
        logger.info(f"[OK   ] : {item_id} .json downloaded")
    else:
        logger.error(f"[FAIL ] : {item_id} .json not downloaded")

    #log error id-rev
    if not file_json.exists() or not file_bmp.exists():
        # allow copy-past of the problematic parts in the log file.
        logger.error(f"id;rev>>>{item_id};{revision}")

    # If server disconnection code 666 then tempo of 15 min.
    if p1.returncode == 666:
        time.sleep(60 * 15)


# maybe we could add some defaults values.
@click.command(
    help="Exemple: batch_rendering -e <excel.xlsx> -o <Path-to-download-output> (Ctrl + C to stop the macro)."
)
@click.option(
    "--excel",
    "-e",
    nargs=1,
    required=True,
    help="Path to excel.",
    type=click.Path(exists=True),
)
@click.option(
    "--user",
    "-u",
    default=user_name,
    show_default=True,
    nargs=1,
    required=False,
    type=str,
    help="User acronym. (Optional if matching the username of your PC.)",
)
@click.password_option(help="Prompt.")
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
    show_default=True,
    required=False,
    type=click.Choice(["ERROR", "INFO"]),
    help="Debug mode for logging file.",
)
@click.option(
    "--single",
    required=False,
    type=int,
    help="Run one single ligne of the excel file.",
)
@click.option(
    "--to-row",
    required=False,
    nargs=1,
    type=int,
    help="",
)
@click.option(
    "--from-row",
    required=False,
    nargs=1,
    type=int,
    help="",
)
@click.option(
    "--between-rows",
    required=False,
    nargs=2,
    type=int,
    help="",
)
def batch_rendering(
    excel, user, password, folder_target, group, role, server, mode, debug_mode, single=False, to_row=False, from_row=False, between_rows=False
):
    """Run SEToolRendering in batch with an excel file."""

    # Ch
    if se_opened_instance() == 0:
        print("No SolidEdge session open.")
        return
    elif se_opened_instance() > 1:
        print("Only one SolidEdge session can be open to run the macro.")
        return


    # Location of the log file.
    path_log = os.path.join(folder_target, "rendering_error.log")
    logger.add(
        path_log,
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
    r = Range(number_rows)

    # Check the rows selected.
    if single:
        rows = r.unique_row(single)
    elif to_row:
        rows = r.to_row(to_row)
    elif from_row:
        rows = r.from_row(from_row)
    elif between_rows:
        rows = r.between_rows(between_rows[0], between_rows[1])
    else:
        rows = range(number_rows)

    try:
        tic = time.perf_counter()
        # For each row run a command line.
        for row in rows:
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
        # Send a single notification
        notifier = notifiers.get_notifier("gmail")
        notifier.notify(message="The macro has finished the process.", **params)
