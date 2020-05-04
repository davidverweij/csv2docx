import csv2docx.csv2docx as c2d
import click



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
    '--destination', '-dest',
    default="output",
    help='The destination to store the files.')
def main(data, template, delimiter, destination):
    c2d.convert(data, template, delimiter, destination)
