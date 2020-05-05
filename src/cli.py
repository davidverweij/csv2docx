import click
from csv2docx import convert


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
    '--name', '-n',
    required=True,
    help='naming scheme for output files.')
@click.option(
    '--path', '-p',
    default="output",
    help='The location to store the files.')
def main(data, template, delimiter, name, path):
	convert(data, template, delimiter, name, path)
