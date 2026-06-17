import subprocess
import pandas as pd
import os

# Optional: define known target repos (better than scanning everything)
TARGET_REPOS = ["UTC_Helper", "SpellChecker"]

# read changed files
with open("changed_files.txt", "r") as f:
    changed_files = f.read().splitlines()

# filter only .py files
changed_py_files = [f for f in changed_files if f.endswith(".py")]

if not changed_py_files:
    print("No Python changes detected. Skipping report generation.")
    exit(0)

# group files by repo
repo_files = {}

for file in changed_py_files:
    for repo in TARGET_REPOS:

        full_path = os.path.join(repo, file)

        if os.path.exists(full_path):
            repo_files.setdefault(repo, [])

            # avoid duplicates
            if full_path not in repo_files[repo]:
                repo_files[repo].append(full_path)

# process each repo separately
for repo, files in repo_files.items():

    excel_file = f"{repo}_PEP8.xlsx"

    with pd.ExcelWriter(excel_file, engine="openpyxl") as writer:

        for file_path in files:

            result = subprocess.run(
                ["python", "-m", "pylint", file_path, "--score=no"],
                capture_output=True,
                text=True
            )

            violations = []

            for line in result.stdout.splitlines():
                violations.append({"Violation": line})

            if not violations:
                violations.append({"Violation": "No issues found"})

            df = pd.DataFrame(violations)

            # unique + safe sheet name
            sheet_name = f"{os.path.basename(file_path).replace('.py','')}"[:31]

            df.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"{excel_file} generated")

# save all report names (optional, not needed for your current YAML)
with open("report_name.txt", "w") as f:
    for repo in repo_files.keys():
        f.write(f"{repo}_PEP8.xlsx\n")
