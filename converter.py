from openpyxl import Workbook

from parser import (
    get_file_date,
    parse_txt_file,
    group_measurements
)

from report_parser import (
    USE_REPORT_BOUNDARY,
    get_selected_report
)

from workbook_builder import (
    create_workbook,
    create_main_header,
    create_information_section,
    create_measurement_table
)

from summary import create_summary


# -----------------------------------
# Convert TXT to Excel
# -----------------------------------

def convert_file(input_file, output_file):
    """
    Convert one TXT weighing file
    into a formatted Excel file.
    """

    # -----------------------------------
    # 1. Get date
    # -----------------------------------

    file_date = get_file_date(
        input_file
    )


    # -----------------------------------
    # 2. Read TXT file
    # -----------------------------------

    if USE_REPORT_BOUNDARY:

        report = get_selected_report(
            input_file
        )

        time_started = report["time_started"]
        measurements = report["measurements"]

    else:

        time_started, measurements = parse_txt_file(
            input_file
        )


    # -----------------------------------
    # 3. Divide measurements into groups
    # -----------------------------------

    groups = group_measurements(
        measurements
    )


    if not groups:

        raise ValueError(
            "No measurement groups were created."
        )


    # -----------------------------------
    # 4. Determine table width
    # -----------------------------------

    total_columns = len(groups) * 2


    # -----------------------------------
    # 5. Create workbook
    # -----------------------------------

    workbook, sheet = create_workbook()


    # -----------------------------------
    # 6. Create main header
    # -----------------------------------

    create_main_header(
        sheet,
        total_columns
    )


    # -----------------------------------
    # 7. Create information section
    # -----------------------------------

    create_information_section(
        sheet,
        total_columns,
        time_started,
        file_date
    )


    # -----------------------------------
    # 8. Create measurement table
    # -----------------------------------

    create_measurement_table(
        sheet,
        groups,
        total_columns
    )


    # -----------------------------------
    # 9. Create final summary
    # -----------------------------------

    create_summary(
        sheet,
        groups,
        total_columns
    )


    # -----------------------------------
    # 10. Save Excel
    # -----------------------------------

    workbook.save(
        output_file
    )


    # -----------------------------------
    # 11. Return information to GUI
    # -----------------------------------

    return {
        "file": output_file,
        "date": file_date,
        "time_started": time_started,
        "measurements": len(measurements),
        "groups": len(groups)
    }