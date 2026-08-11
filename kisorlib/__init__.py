import warnings

# Suppress openpyxl Data Validation extension user warning
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
