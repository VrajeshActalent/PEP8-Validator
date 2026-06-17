import subprocess
import pandas as pd
import os

excel_file = "pylint_multi_report.xlsx"

with pd.ExcelWriter(excel_file, engine="openpyxl") as writer:
    
    for folder in os.listdir("."):
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

                sheet_name = file.replace(".py", "")[:31]

                df.to_excel(writer, sheet_name=sheet_name, index=False)

# save report name
with open("report_name.txt", "w") as f:
    f.write(excel_file)

print("Multi-file report generated")
