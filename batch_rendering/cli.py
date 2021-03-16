"""
Only works in mode TC.
"""

import subprocess
import time
import os, sys

import click
import pandas as pd
import datetime


def render_one_item(items: list, user, password, group, role, server, folder_target, mode="TC"):
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


@click.command()
@click.argument("excel", nargs=1)
@click.argument("user", nargs=1)
@click.argument("password", nargs=1)
@click.argument("group", nargs=1)
@click.argument("role", nargs=1)
@click.argument("server", nargs=1)
@click.argument("folder_target", nargs=1)
def batch_rendering(excel, user, password,  group, role, server, folder_target):
    """Run SEToolRendering in batch with an excel file."""

    df = pd.read_excel(
        excel,
        index_col=None,
        dtype={
            "item ID (-i)": str,
            "revision (-v)": str,
        },
    )
    # Define the shape of the excel file.
    number_rows, number_columns = df.shape

    try:
        tic = time.perf_counter()
        # For each row run a command line.
        for row in range(number_rows):
            items: list = [df.iat[row, column] for column in range(number_columns)]
            render_one_item(items)
            print(f"[OK  ] iteration: [{row+1}/{number_rows}]")

    except subprocess.CalledProcessError as c:
        with open("log_rendering_errors.txt", "a") as err:
            err.write(f"[FAIL] Excel line: {row+1}")
            print(f"[FAIL] iteration: [{row+1}/{number_rows}]")

    finally:
        toc = time.perf_counter()
        seconds = toc - tic
        print(f"Time to completion: {datetime.timedelta(seconds=seconds)}")
