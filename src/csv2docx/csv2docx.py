import csv
from pathlib import Path
import warnings

from mailmerge import MailMerge


def create_output_folder(output_path: str) -> Path:
    """Creates a path to store output data if it does not exists.
    Args:
        path: the path from user in any format (relative, absolute, etc.)
    Returns:
        A path to store output data.
    """

    path = Path(output_path)
    if not path.exists():
        path.mkdir(parents=True, exist_ok=True)
    return path


def create_unique_name(filename: str, path: Path) -> Path:
    """Creates a unique filename for specified path.
    Args:
        filename: the name of file to create
        path: the path where the file is stored
    Returns:
        An absolute path with directory.
    """
    filename = filename.strip() + ".docx"
    filepath = path / filename
    if filepath.exists():
        # Count available files with same name
        counter = len(list(path.glob(f"{filepath.stem}*docx"))) + 1
        filepath = path / f"{filepath.stem}_{counter}.docx"
    return filepath


def generate_docx(data: dict, template: str, mergefields: set) -> MailMerge:
    """Generates a single docx
    Args:
        data: a dictionary representing a single .csv row
        template: the name of the template .docx
        mergefields: a set of fields to be filled by the data
    Returns:
        A MailMerge (.docx) object with filled values
    """

    # Must create a new MailMerge for each file
    docx = MailMerge(template)
    fields = {key: data[key] for key in mergefields}
    docx.merge_templates([fields], separator="page_break")
    return docx


def convert(
    data: str, template: str, name: str, path: str = "output", delimiter: str = ";"
) -> bool:

    try:
        csvfile = open(data, "rt")
    except FileNotFoundError:
        raise FileNotFoundError(f"The data .csv file '{data}' could not be found")
    else:
        with csvfile:
            csvdict = csv.DictReader(csvfile, delimiter=delimiter)
            csv_headers = csvdict.fieldnames

            try:
                docx = MailMerge(template)
            except FileNotFoundError:
                raise FileNotFoundError(
                    f"The template .docx file '{template}' could not be found"
                )

            docx_mergefields = docx.get_merge_fields()

            if csv_headers and name not in csv_headers:
                if len(csv_headers) == 1 and len(docx_mergefields) > 1:
                    warnings.warn(
                        f"Only one .csv column could be found, but multiple"
                        f" mergefields, potentially the .csv has a different"
                        f" delimiter than '{delimiter}'",
                        UserWarning,
                    )
                raise ValueError(
                    f"The column {name} could not found in the .csv header "
                    f"to be used in the naming scheme."
                )

            # see if all fields are accounted for in the .csv header
            column_in_data = set(docx_mergefields) - set(csv_headers)
            if len(column_in_data) > 0:
                raise KeyError(
                    f"{column_in_data} are mailmerge fields in the template, "
                    f"but missing in the .csv header"
                )

            output_path = create_output_folder(path)

            for row in csvdict:
                # check if at least one key is a non-empty string
                if any(row.values()):
                    docx = generate_docx(
                        data=row, template=template, mergefields=docx_mergefields
                    )
                    filename = create_unique_name(row[name], output_path)
                    docx.write(filename)

            return True
