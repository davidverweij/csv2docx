import click
from . import csv2docx


@click.command()
@click.option(
    '--data', '-c',
    required=True,
    help='Path to the .csv data file to be used')
@click.option(
    '--template', '-t',
    required=True,
    help='Path to the .docx template file to be used')
@click.option(
    '--delimiter', '-d',
    default=";",
    help='delimiter used in your csv. Default is \';\'')
@click.option(
    '--name', '-n', 'column_name',
    help='Naming scheme for output files.')
def main(data, template, delimiter, column_name):
    csv2docx.convert(data, template, delimiter, column_name)
