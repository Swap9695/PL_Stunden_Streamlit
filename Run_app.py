import subprocess
import webbrowser

# Start Streamlit
subprocess.Popen(["streamlit", "run", "PL_Stunden_Streamlit.py"])

# Open browser automatically
webbrowser.open("http://localhost:8501")
