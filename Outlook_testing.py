import win32com.client

try:
    outlook = win32com.client.gencache.EnsureDispatch('Outlook.Application')
    print("Outlook COM loaded successfully.")
except Exception as e:
    print(f"Failed to load Outlook COM: {e}")
