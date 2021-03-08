import pandas as pd
import subprocess
import click


def render_one_item(commands: list, mode: str):
    if mode == "SE":
        mode, source, source_file, folder_target, file_target = (
            commands[0],
            commands[1],
            commands[2],
            commands[3],
            commands[4],
        )
        subprocess.run(
            f"SEToolRender -m {mode} -f {source} -i {source_file} -o {folder_target} -q {file_target}",
            shell=True,
            check=True,
        )
    elif mode == "TC":
        mode, item_id, revision, user, password, group, role, server, folder_target = (
            commands[0],
            commands[1],
            commands[2],
            commands[3],
            commands[4],
            commands[5],
            commands[6],
            commands[7],
        )
        subprocess.run(
            f"SEToolRender -m {mode} -i {item_id} -f {revision} -u {user} -p {password} -g {group} -r {role} -s {server} -o {folder_target}",
            shell=True,
            check=True,
        )
    else:
        raise NotImplementedError

@click.command()
@click.argument('excel')
def batch_rendering(excel):
    """Run SEToolRendering in batch with an excel file."""
    df = pd.read_excel(
        # r"C:\Users\recs\Desktop\RenderingTool\template_rendering.xlsx",
        excel,
        index_col=None,
        dtype={
            "mode (-m)": str,
            "folder source (-f)": str,
            "file source (-i)": str,
            "folder target (-o)": str,
            "file target (-q)": str,
        },
    )

    # Define the shape of the excel file.
    number_rows, number_columns = df.shape

    # For each row run a command line.
    for row in range(number_rows):
        commands: list = [df.iat[row, column] for column in range(number_columns)]
        render_one_item(commands, mode=commands[0])
