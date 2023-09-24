# Add hidden imports for modules that PyInstaller couldn't automatically detect.
hiddenimports = [
    'pandas',
    'selenium',
    'selenium.webdriver.chrome.service',
    'selenium.webdriver.chrome.options',
    'selenium.webdriver.support',
    'selenium.webdriver.support.expected_conditions',
    'selenium.webdriver.support.ui',
    'selenium.webdriver.common.by',
    'webdriver_manager',
    'sendgrid',
    'sendgrid.helpers.mail',
    'embassy',
]

# Specify additional files or directories to include (e.g., data files).
datas = [
    ('user_details.xlsx', '.'),  # Include the Excel file in the bundle.
    ('config.ini', '.'),         # Include the config file in the bundle.
]

# Create the EXE
exe = EXE(
    # Path to your Python script.
    script='visa_v.5.py',
    # The name of the generated executable.
    name='visa_sche',
    # Additional options, if needed.
    # ...
)

# Binarize libraries, if necessary.
binaries = [
    # ('path/to/library.dll', '.'),  # Example: Include a DLL file.
]

# Include any necessary packages and modules.
coll = COLLECT(
    binaries=binaries,
    datas=datas,
    exe=exe,
    # Additional options, if needed.
    # ...
)

# Return the COLLECT configuration.
