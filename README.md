# KashMap Migration Suite

Unified multi-platform BI Migration Suite powered by Streamlit.

## Project Structure
```
KashMap/
|-- launcher.py
|-- requirements.txt
|-- migrations/
|   |-- webfocus_powerbi/
|   |   `-- app.py
|   |-- webfocus_cognos/
|   |   `-- app.py
|   |-- webfocus_tableau/
|   |   `-- app.py
|   |-- webfocus_pyramid/
|   |   `-- app.py
|   |-- cognos_powerbi/
|   |   `-- app.py
|   `-- tableau_powerbi/
|       `-- app.py
`-- pages/
    |-- webfocus_powerbi.py
    |-- webfocus_cognos.py
    |-- webfocus_tableau.py
    |-- webfocus_pyramid.py
    |-- cognos_powerbi.py
    `-- tableau_powerbi.py
```

## Running the Application
```bash
streamlit run launcher.py
```
