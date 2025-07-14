import os
import sys

# Add the ccpm_python directory to the python path
sys.path.append(os.path.abspath("ccpm_python"))

# Change the current working directory to the ccpm_python directory
os.chdir("ccpm_python")

# Run the main.py script
os.system("python3 main.py")
