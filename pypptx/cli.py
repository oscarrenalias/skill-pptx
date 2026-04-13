import json
import sys
from typing import Callable

import click

import pypptx


def output_result(data: dict, plain: bool, plain_fn: Callable[[dict], str]) -> None:
    """Write data to stdout as JSON or as plain text via plain_fn."""
    if plain:
        sys.stdout.write(plain_fn(data) + "\n")
    else:
        sys.stdout.write(json.dumps(data) + "\n")


@click.group()
@click.version_option(version=pypptx.__version__, prog_name="pypptx")
def cli() -> None:
    """pypptx — PowerPoint manipulation toolkit."""


@cli.group()
def slide() -> None:
    """Commands for working with slides."""
