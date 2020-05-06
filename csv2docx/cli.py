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
    help='Delimiter used in your csv. Default is \';\'')
@click.option(
    '--name', '-n',
    required=True,
    help='Naming scheme for output files.')
@click.option(
    '--path', '-p',
    default="output",
    help='The location to store the files.')
def main(data, template, name, delimiter, path):
    csv2docx.convert(data, template, name, path, delimiter)
