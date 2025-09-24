import subprocess

scripts = [
    "sales_orders_report.py",
    # Add other report scripts here
]

for s in scripts:
    print(f"Running {s}...")
    subprocess.run(["python", s], check=True)
