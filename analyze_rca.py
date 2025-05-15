
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import re

# Load the Excel file
input_path = "sample_incidents.xlsx"
df = pd.read_excel(input_path)

# Normalize the Short Description
def clean_text(text):
    return re.sub(r'\W+', ' ', str(text).lower().strip())

df["Clean Description"] = df["Short Description"].apply(clean_text)

# Detect Recurring Issues
desc_counts = df["Clean Description"].value_counts()
df["Is Recurring"] = df["Clean Description"].apply(lambda x: "Yes" if desc_counts[x] > 1 else "No")

# Generate bar chart of top recurring issues
plt.figure(figsize=(10, 6))
top_repeats = desc_counts[desc_counts > 1]
sns.barplot(x=top_repeats.values, y=top_repeats.index)
plt.title("Top Recurring Issues (Short Descriptions)")
plt.xlabel("Count")
plt.ylabel("Issue Summary")
plt.tight_layout()
chart_path = "RCA_Recurring_Issues_Chart.png"
plt.savefig(chart_path)
plt.close()

# Save updated report
output_path = "RCA_Weekly_Report.xlsx"
df.to_excel(output_path, index=False)

print("Analysis complete. Files saved:")
print(f"- {output_path}")
print(f"- {chart_path}")
