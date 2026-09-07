import re


# -----------------------------------
# Report boundary settings
# -----------------------------------

USE_REPORT_BOUNDARY = True

REPORT_MODE = "first_disconnected"


# -----------------------------------
# Parse reports from TXT file
# -----------------------------------

def parse_reports(input_file):
    """
    Parse the TXT file into report sections.

    A report starts when a Connected event is found
    and ends at the next Disconnected event.

    Connection lost does NOT end a report.

    Returns:
        [
            {
                "time_started": "...",
                "end_time": "...",
                "measurements": [...]
            },
            ...
        ]
    """

    reports = []

    current_report = None

    with open(
        input_file,
        "r",
        encoding="utf-8"
    ) as file:

        for line in file:

            # -----------------------------------
            # Connected
            # -----------------------------------

            connected_match = re.search(
                r"(\d{2}:\d{2}:\d{2}\.\d{3})\s+Connected",
                line
            )

            if connected_match:

                # Start a report only if one
                # is not already active.
                if current_report is None:

                    current_report = {
                        "time_started": connected_match.group(1),
                        "end_time": None,
                        "measurements": []
                    }

                continue


            # -----------------------------------
            # Disconnected
            # -----------------------------------

            disconnected_match = re.search(
                r"(\d{2}:\d{2}:\d{2}\.\d{3})\s+Disconnected",
                line
            )

            if disconnected_match:

                # Ignore Disconnected events that
                # occur before a valid report starts.
                if current_report is None:
                    continue

                current_report["end_time"] = (
                    disconnected_match.group(1)
                )

                # Only keep reports that actually
                # contain measurements.
                if current_report["measurements"]:

                    reports.append(
                        current_report
                    )

                current_report = None

                continue


            # -----------------------------------
            # Weight measurement
            # -----------------------------------

            weight_match = re.search(
                r"(\d{2}:\d{2}:\d{2}\.\d{3})\s+([\d.]+)kg",
                line
            )

            if weight_match:

                # Ignore measurements that occur
                # before a valid Connected event.
                if current_report is None:
                    continue

                current_report["measurements"].append({
                    "time": weight_match.group(1),
                    "weight": float(
                        weight_match.group(2)
                    )
                })


    # -----------------------------------
    # Handle unfinished final report
    # -----------------------------------

    if current_report is not None:

        if current_report["measurements"]:

            reports.append(
                current_report
            )


    return reports


# -----------------------------------
# Select report
# -----------------------------------

def get_selected_report(input_file):
    """
    Return the report selected by REPORT_MODE.
    """

    reports = parse_reports(
        input_file
    )

    if not reports:

        raise ValueError(
            "No valid reports were found."
        )


    # -----------------------------------
    # First report
    # -----------------------------------

    if REPORT_MODE == "first_disconnected":

        return reports[0]


    # -----------------------------------
    # Future modes
    # -----------------------------------

    if REPORT_MODE == "all_reports":

        return reports


    raise ValueError(
        f"Unknown REPORT_MODE: {REPORT_MODE}"
    )