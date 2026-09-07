import os
import re
from datetime import datetime


# -----------------------------------
# Get date from filename
# -----------------------------------

def get_file_date(input_file):
    """
    Extract the date from a filename such as:

    serial_20260823_1234.txt

    Returns:
        08/23/2026
    """

    filename = os.path.basename(input_file)

    filename_match = re.search(
        r"serial_(\d{8})_\d+\.txt$",
        filename
    )

    if not filename_match:
        raise ValueError(
            "Could not determine the date from the filename."
        )

    date_text = filename_match.group(1)

    return datetime.strptime(
        date_text,
        "%Y%m%d"
    ).strftime("%m/%d/%Y")


# -----------------------------------
# Read and parse TXT
# -----------------------------------

def parse_txt_file(input_file):
    """
    Read the TXT weighing file.

    Returns:
        time_started
        measurements
    """

    measurements = []

    time_started = None

    with open(
        input_file,
        "r",
        encoding="utf-8"
    ) as file:

        for line in file:

            # -----------------------------------
            # Connected time
            # -----------------------------------

            connected_match = re.search(
                r"(\d{2}:\d{2}:\d{2}\.\d{3})\s+Connected",
                line
            )

            if connected_match and time_started is None:

                time_started = connected_match.group(1)


            # -----------------------------------
            # Weight measurement
            # -----------------------------------

            weight_match = re.search(
                r"(\d{2}:\d{2}:\d{2}\.\d{3})\s+([\d.]+)kg",
                line
            )

            if weight_match:

                time = weight_match.group(1)

                weight = float(
                    weight_match.group(2)
                )

                measurements.append({
                    "time": time,
                    "weight": weight
                })


    # -----------------------------------
    # Check measurements
    # -----------------------------------

    if not measurements:

        raise ValueError(
            "No weight measurements were found."
        )


    return time_started, measurements


# -----------------------------------
# Divide measurements into groups
# -----------------------------------

def group_measurements(
    measurements,
    group_size=10
):
    """
    Divide measurements into groups.

    Default:
        10 measurements per group.
    """

    return [
        measurements[i:i + group_size]
        for i in range(
            0,
            len(measurements),
            group_size
        )
    ]