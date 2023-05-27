import matplotlib.pyplot as plt
from openpyxl import load_workbook

# the point of this script is to extract the latest data from transelectrica.ro 's SEN Graph section
# after running script_Download_slsx
# we should have the file Grafic-SEN.xlsx downloaded, which contains data about energy consumption and production
# the following lines retrieve it
excel_file_path = r"C:\Users\paula\Downloads\Grafic_SEN.xlsx"
workbook = load_workbook(excel_file_path)
worksheet = workbook.active

# after retrieving the file, we extract the values
# since we want to extract the latest data
# and the first row of the table contains the name of the attributes
# we are extracting values from the second row of the table
date_time = worksheet['A2'].value
consumption = worksheet['B2'].value
hourly_average = worksheet['C2'].value
production = worksheet['D2'].value
coal = worksheet['E2'].value
hydrocarbons = worksheet['F2'].value
hydro = worksheet['G2'].value
nuclear = worksheet['H2'].value
wind = worksheet['I2'].value
photovoltaic = worksheet['J2'].value
biomass = worksheet['K2'].value
balance = worksheet['L2'].value

# generating the text to be included in the html file
sentence = f"According to transelectrica.ro, at {date_time},<br><br>" \
           f"The energy Consumption was {consumption} MW, <br> The Hourly Average for Consumption was {hourly_average} MW, <br>" \
           f"the Production was {production} MW, <br> from the following sources: <br><br> Coal = {coal} MW, <br>" \
           f"Hydrocarbons = {hydrocarbons} MW, <br> Hydro = {hydro} MW, <br> Nuclear = {nuclear} MW, <br>" \
           f"Wind = {wind} MW, <br> Photovoltaic = {photovoltaic} MW, <br> Biomass = {biomass} MW.<br><br>" \
           f"The Consumption - Production Balance was {balance} MW."

# generating the bar chart
# it displays the production value for every type of energy source
categories = ['Coal', 'Hydrocarbons', 'Hydro', 'Nuclear', 'Wind', 'Photovoltaic', 'Biomass']
values = [coal, hydrocarbons, hydro, nuclear, wind, photovoltaic, biomass]

plt.bar(categories, values, width = 0.4)
plt.xlabel('ENERGY SOURCES')
plt.ylabel('MW')
plt.title('ENERGY PRODUCTION BY SOURCE')

# saving the bar char as a png
# which will later be included in the html
chart_path = r'C:\Users\paula\source\repos\data-extraction-to-html\result\energy_chart.png'
plt.savefig(chart_path)
plt.close()

# creating the html file
# adding the text and png
email_content = f"""<html>
<head></head>
<body>
<p>{sentence}</p>
<img src="{chart_path}" alt="Energy Production Chart">
</body>
</html>"""

# saving the html file in my desired directory
# every time we run this script, a new file will be created with up to date information
email_file_path = r'C:\Users\paula\source\repos\data-extraction-to-html\result\email_content.html'
with open(email_file_path, 'w') as email_file:
    email_file.write(email_content)

print('Email content saved as HTML file.')