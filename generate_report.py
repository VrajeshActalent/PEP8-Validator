import subprocess
import pandas as pd
import os

# read changed files
with open("changed_files.txt", "r") as f:
    changed_files = f.read().splitlines()

# filter only .py files
changed_py_files = [f for f in changed_files if f.endswith(".py")]

if not changed_py_files:
    print("No Python changes detected. Skipping report generation.")
    exit(0)

# get repo name dynamically
repo_name = os.getcwd().split("/")[-1]
excel_file = f"{repo_name}_PEP8.xlsx"

with pd.ExcelWriter(excel_file, engine="openpyxl") as writer:

    for file_path in changed_py_files:

        if not os.path.exists(file_path):
            continue

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

        sheet_name = os.path.basename(file_path).replace(".py", "")[:31]

        df.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"{excel_file} generated")
