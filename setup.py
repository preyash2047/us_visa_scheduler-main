from cx_Freeze import setup, Executable

base = None

executables = [Executable("visa_v.5.py", base=base)]

setup(
    name="Visa Scheduler",
    version="1.0",
    description="Application to schedule the visa.",
    executables=executables
)
