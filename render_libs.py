import pandas as pd

# Load the Excel file
file_path = 'JellyFin_Server.xlsx'
excel_data = pd.ExcelFile(file_path)

index_html = """
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <meta name="description" content="My Library for the home JellyFin Server">
        <title>Personal Archive</title>
        <link rel="icon" type="image/x-icon" href="./assets/favicon.svg">
        <link rel="stylesheet" href="./assets/style.css">
    </head>
    
    <body>
        <br>
        <h1 style="margin-bottom: 5px;">Personal Archive
        </h1>
        
        <hr><br><ul>
"""


# Iterate through each sheet in the Excel file
for sheet_name in excel_data.sheet_names:
    
    print(f"Processing sheet: {sheet_name}")
    
    df = pd.read_excel(excel_data, sheet_name=sheet_name)
    
    file_name = f"{sheet_name}.html"
    
    html = f"""
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <meta name="description" content="{sheet_name}">
        <title>{sheet_name}</title>
        <link rel="icon" type="image/x-icon" href="./assets/favicon.svg">
        <link rel="stylesheet" href="./assets/style.css">
    </head>
    
    <body>
        <br>
        <h1 style="margin-bottom: 5px;">{sheet_name}
            <a href="./index.html">
                <span style="font-size: 20px; float: right;">⇦ Back to Index</span>
            </a>
        </h1>
        
        <hr><br>
    """
    
    index_html += f'\n<li><a href="{file_name}">{sheet_name}</a></li>'
    
    # Organize by year
    df = df.sort_values(by="Year")
    
    # Group by the "Category" column
    grouped = df.groupby("Category")
        
    # Iterate through each category and its rows
    for category, group in grouped:
        
        html += f"    <h2>{category}</h2><ul>"
        
        for index, row in group.iterrows():
            
            year, director, author = "", "", ""
            
            # If Year is filled
            if not pd.isna(row["Year"]):
                year = f'<code>[Year: {int(row["Year"])}]</code>'
                
            # Check if collumn "Author" exists
            if "Director" in row.index:
                if not pd.isna(row["Director"]):
                    author = f'<code>[Director: {row["Director"]}]</code>'
        
            # Check if collumn "Author" exists
            if "Author" in row.index:
                if not pd.isna(row["Author"]):
                    author = f'<code>[Author: {row["Author"]}]</code>'



            
            print(f"Adding {row['Title']} to {category}...")
            html += f'\n        <li><a href="{row["Magnetic Link"]}" target="_blank">{row["Title"]} {year} {director} {author}</a></li>'
        
        html += "\n</ul>"
    html += "\n</body></html>"
    
    # Save the HTML file with UTF-8 encoding
    with open(file_name, 'w', encoding='utf-8') as f:
        f.write(html)
        
index_html += "\n</ul></body></html>"

# SAVE THE INDEX HTML
with open('index.html', 'w', encoding='utf-8') as f:
    f.write(index_html)

print("HTML files have been generated.")
