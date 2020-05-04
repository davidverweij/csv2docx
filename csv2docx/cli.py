import click
import csv2docx.csv2docx as c2d


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
    help='naming scheme for output files.')
def main(data, template, delimiter, name):
	c2d.convert(data, template, delimiter, name)
