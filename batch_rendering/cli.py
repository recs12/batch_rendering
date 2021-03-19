"""
Only works in mode TC.
"""

import subprocess
import time
import os

import click
import pandas as pd
import datetime
from loguru import logger
import pathlib


def render_one_item(items, user, password, folder_target, group, role, server, mode):
    print(f"[INFO] TC - Part Number: {items[1]}")
    item_id, revision = items[0], items[1]
    subprocess.run(
        f"SEToolRender -m {mode} -i {item_id} -v {revision} -u {user} -p {password} -g {group} -r {role} -s {server} -o {folder_target}",
        shell=True,
        check=True,
        capture_output=True,
        text=True,
    )
    # check the presence of the files (json, bmp)
    bmp = os.path.join(folder_target, f"{item_id}-Rev-{revision}.bmp")
    json = os.path.join(folder_target, f"{item_id}-Rev-{revision}.json")
    file_bmp = pathlib.Path(bmp)
    file_json = pathlib.Path(json)
    if file_bmp.exists() and file_json.exists():
        logger.info(f"[OK  ] : {item_id} Files downloaded")
    else:
        logger.error(f"[FAIL] : {item_id} Files not downloaded")


# maybe we could add some defaults values.
@click.command(
    help="excel.xlsx -u <user> -p <password> -o <target-folder> <group>* <role>* <server>*"
)
@click.argument("excel", nargs=1, required=True, type=click.Path(exists=True))
@click.option("--user", "-u", nargs=1, required=True, type=str, help="Your acronym.")
@click.option(
    "--password",
    "-p",
    prompt=True,
    confirmation_prompt=True,
    hide_input=True,
    required=True,
    help="Your password.",
)
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
    default="Engineering", show_default=True,
    nargs=1,
    required=False,
    type=click.Choice(['Engineering']),
    help="Group optional.",
)
@click.option(
    "--role",
    "-r",
    default="Designer", show_default=True,
    nargs=1,
    required=False,
    type=click.Choice(['Designer']),
    help="Role optional.",
)
@click.option(
    "--server",
    "-s",
    default="TC_PROD", show_default=True,
    nargs=1,
    required=False,
    type=click.Choice(['TC_PROD']),
    help="Server optional.",
)
@click.option(
    "--mode",
    "-m",
    default="TC", show_default=True,
    nargs=1,
    required=False,
    type=click.Choice(['TC']),
    help="TC is the only mode enabled.",
)
def batch_rendering(
    excel,
    user,
    password,
    folder_target,
    group,
    role,
    server,
    mode,
):
    """Run SEToolRendering in batch with an excel file."""

    logger.add(
        "report_rendering.log",
        format="{time:YYYY-MM-DD at HH:mm:ss} | {level} | {message}",
        level="INFO",
    )

    df = pd.read_excel(
        excel, index_col=None, dtype={"item ID (-i)": str, "revision (-v)": str}
    )
    # Define the shape of the excel file.
    number_rows, number_columns = df.shape

    try:
        tic = time.perf_counter()
        # For each row run a command line.
        for row in range(number_rows):
            items: list = [df.iat[row, column] for column in range(number_columns)]
            render_one_item(items, user, password, folder_target, group, role, server, mode)
            logger.info(f"[ITERATION] {row+1}/{number_rows}")

    except subprocess.CalledProcessError:
        logger.exception({row + 1})

    finally:
        toc = time.perf_counter()
        seconds = toc - tic
        logger.info(f"Time to completion: {datetime.timedelta(seconds=seconds)}")
