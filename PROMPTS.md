**Prompt 1:**
Help me to build a web app, which can be an IC50 calculator for 3CLpro inhibitors. I want the app to allow users to upload a CSV file, which contains two columns: concentration and fluorescence value. The app should calculate the relative enzyme activity, and if there is background blank value, subtract it. Then, fit a sigmoid curve to the data. Plot a scatter plot with the curve on a logarithmic x-axis. Finally, show the calculated IC50 value clearly.

**Prompt 2:**
I need you to fix 2 issues.
1. I can't upload CSV files that aren't named with ".csv" extension
2. The text on the website is too small and hard to read

(I updated the app to accept any CSV files that meet the format requirements and made the text visually easier to read.)

**Prompt 3:**
1. Also accpt excel files 
2. Add "AI Comment", which is a short, brief AI comment on the data, under the result 
3. Allow users to upload more than 1 files and show all curves

(I modified the app to accept multiple CSV or Excel files and generate an AI comment of the results.)

**Prompt 4:**
I can't see where to remove uploaded files. You can add a small trash bin button to the upper right corner of each data box

(I added a function that lets users remove uploaded data.)