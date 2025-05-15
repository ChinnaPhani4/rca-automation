
import os
import win32com.client as win32

# File paths
report_file = "RCA_Weekly_Report.xlsx"
chart_file = "RCA_Recurring_Issues_Chart.png"

# Create the email
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'chinnaphani4@outlook.com'  # <- Replace with real recipient
mail.CC = ''                     # <- Optional: Add CCs
mail.Subject = 'Weekly RCA Report'
mail.Body = (
    'Hi Team,\n\n'
    'Please find attached the RCA summary report and recurring issue chart for this week.\n\n'
    'Regards,\nPhani'
)

# Attach files
mail.Attachments.Add(os.path.abspath(report_file))
mail.Attachments.Add(os.path.abspath(chart_file))

# Send the email (or use mail.Display() to preview)
#mail.Send()
mail.Display()
# mail.Display()  # Uncomment to preview instead of send
print("Email sent successfully.")
