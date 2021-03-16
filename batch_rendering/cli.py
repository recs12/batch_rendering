"""
Only works in mode TC.
"""

import subprocess
import time
import os, sys

import click
import pandas as pd
import datetime


def render_one_item(commands: list, mode: str):
        print(f"[INFO] TC - Part Number: {commands[1]}")
        mode, item_id, revision, user, password, group, role, server, folder_target = (
            "TC",
            commands[0],
            commands[1],
            "recs",
            "recs",
            "",
            "",
            "",
            "",
        )
        subprocess.run(
            f"SEToolRender -m {mode} -i {item_id} -f {revision} -u {user} -p {password} -g {group} -r {role} -s {server} -o {folder_target}",
            shell=True,
            check=True,
            capture_output=True,
            text=True,
        )
        # check the presence of the files (json, bmp)



@click.command()
@click.argument("excel")
@click.argument("excel")
@click.argument("excel")
@click.argument("excel")
def batch_rendering(excel):
    """Run SEToolRendering in batch with an excel file."""
    df = pd.read_excel(
        excel,
        index_col=None,
        dtype={
            "mode (-m)": str,
            "folder target (-o)": str,
            "item ID (-i)": str,
            "revision (-v)": str,
            "user (-u)": str,
            "password (-p)": str,
            "group (-g)": str,
            "role (-r)": str,
            "server (-s)": str,
        },
    )

    # Define the shape of the excel file.
    number_rows, number_columns = df.shape

    try:
        tic = time.perf_counter()
        # For each row run a command line.
        for row in range(number_rows):
            commands: list = [df.iat[row, column] for column in range(number_columns)]
            render_one_item(commands, mode=commands[0])
            print(f"[OK  ] iteration: [{row+1}/{number_rows}]")

    except subprocess.CalledProcessError as c:
        with open("log_rendering_errors.txt", "a") as err:
            err.write(f"[FAIL] Excel line: {row+1}")
            print(f"[FAIL] iteration: [{row+1}/{number_rows}]")

    finally:
        toc = time.perf_counter()
        seconds = toc - tic
        print(f"Time to completion: {datetime.timedelta(seconds=seconds)}")
