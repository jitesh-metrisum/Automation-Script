import subprocess

scripts = [ "sales_orders_report.py"]

for s in scripts:
    print(f"Running {s}...")
    subprocess.run(["python", s], check=True)
    print(f"{s} finished.")
