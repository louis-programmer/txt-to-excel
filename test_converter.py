from converter import convert_file


input_file = "serial_20260719_172801.txt"
output_file = "test_output.xlsx"


result = convert_file(
    input_file,
    output_file
)


print("Excel file created:", result["file"])
print("Date:", result["date"])
print("Time started:", result["time_started"])
print("Measurements:", result["measurements"])
print("Groups:", result["groups"])