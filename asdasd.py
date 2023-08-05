import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

# Sample survey data (replace this with your actual survey data)
data = {
    'Gender': ['Male', 'Female', 'Male', 'Male', 'Female', 'Male', 'Female', 'Male', 'Female', 'Female'],
    'Age Group': ['18-25', '26-35', '18-25', '36-45', '26-35', '26-35', '36-45', '26-35', '18-25', '36-45'],
    'Answer 1': ['Yes', 'No', 'No', 'Yes', 'Yes', 'Yes', 'No', 'Yes', 'Yes', 'No'],
    'Answer 2': ['No', 'Yes', 'Yes', 'No', 'No', 'Yes', 'No', 'Yes', 'No', 'No'],
}

# Create a DataFrame from the sample data
df = pd.DataFrame(data)

# Create a crosstab of 'Gender' and 'Answer 1'
cross_tab = pd.crosstab(df['Gender'], df['Answer 1'])

# Use Seaborn heatmap to visualize the crosstab
plt.figure(figsize=(8, 6))
sns.heatmap(cross_tab, annot=True, fmt='d', cmap='YlGnBu', cbar=True)
plt.title('Crosstab: Gender vs. Answer 1')
plt.xlabel('Answer 1')
plt.ylabel('Gender')
plt.show()
