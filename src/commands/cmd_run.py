"""Standard library"""
import os
import csv
from datetime import date

"""Third party modules"""
import click
from tabulate import tabulate
import xlwt
from xlwt import Workbook

"""Internal application modules"""
from src.main import pass_environment


@click.command("run", short_help="Run the sort/format on specified excel spreadsheet.")
@click.option(
    "--path", "-p", required=True, help="Path to excel spreadsheet.",
)
@pass_environment
def cli(ctx, path):
    """Run the sorting and formatting for the excel spreadsheet\n
       Ex. rpcli run --path path/to/file.xlsx"""

    # Ensure path provided to excel spreadsheet exists
    filepath = os.path.join(path)
    if not os.path.exists(filepath):
        raise click.UsageError("Path to excel spread sheet does not exist: '%s'" % path)

    # Create list of names of people from studio
    personnel_list = []

    # Iterate over studio personnel to create list
    with open("data/studio-personnel.csv", newline="", encoding="utf-8-sig") as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            personnel_list.append(row.get("Name"))

    # create excel workbook to export data to
    wb = Workbook()
    sheet = wb.add_sheet("Sheet 1")
    style = xlwt.easyxf("font: bold 1")

    # create excel sheet headers
    sheet.write(0, 0, "Enterprise ID", style)
    sheet.write(0, 1, "Fiscal Year", style)
    sheet.write(0, 2, "Hours Date", style)
    sheet.write(0, 3, "Pay Period Ending", style)
    sheet.write(0, 4, "Account Name", style)
    sheet.write(0, 5, "Project Number", style)
    sheet.write(0, 6, "Project Name", style)
    sheet.write(0, 7, "Entered Hours", style)

    # row counter
    i = 1

    # Iterate over system data to extract names from personnel list
    with open(filepath, newline="", encoding="utf-8-sig") as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            if row.get("Enterprise ID") in personnel_list:
                sheet.write(i, 0, row.get("Enterprise ID"))
                sheet.write(i, 1, row.get("Fiscal Year"))
                sheet.write(i, 2, row.get("Hours Date"))
                sheet.write(i, 3, row.get("Pay Period Ending"))
                sheet.write(i, 4, row.get("Account Name"))
                sheet.write(i, 5, row.get("Project Number"))
                sheet.write(i, 6, row.get("Project Name"))
                sheet.write(i, 7, row.get("Entered Hours"))
                i += 1

    # Get data for report name
    today = date.today()

    # After all data has been written to excel sheet save the workbook
    excelName = "Report " + str(today) + ".xls"
    wb.save("reports/" + excelName)
