import subprocess
import pandas as pd
import os

EXCLUDE_DIRS = ["PEP8_VALIDATOR", ".git", "__pycache__"]

excel_file = "pylint_multi_report.xlsx"

with pd.ExcelWriter(excel_file, engine="openpyxl") as writer:

    for folder in os.listdir("."):

        # only process directories and skip unwanted ones
        if os.path.isdir(folder) and folder not in EXCLUDE_DIRS:

            for root, dirs, files in os.walk(folder):

                for file in files:
                    if file.endswith(".py"):

                        file_path = os.path.join(root, file)

                        result = subprocess.run(
                            ["python", "-m", "pylint", file_path],
                            capture_output=True,
                            text=True
                        )

                        violations = []

                        for line in result.stdout.splitlines():
                            violations.append({"Violation": line})

                        if not violations:
                            violations.append({"Violation": "No issues found"})

                        df = pd.DataFrame(violations)

                        sheet_name = f"{folder}_{file.replace('.py','')}"[:31]

                        df.to_excel(writer, sheet_name=sheet_name, index=False)

# save report name
with open("report_name.txt", "w") as f:
    f.write(excel_file)

print("Multi-repo report generated successfully")
