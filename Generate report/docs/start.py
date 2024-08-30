import subprocess
 
subprocess.run(
    ".venv\\Scripts\\activate && python generate_report.py", shell=True
)