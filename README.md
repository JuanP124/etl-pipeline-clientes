# ETL Pipeline for Client Data

A Python-based Extract, Transform, and Load (ETL) pipeline designed to ingest, clean, and structure raw client metrics before loading them into a PostgreSQL database.

🚀 Features
- **Extraction:** Generates and ingests raw synthetic datasets.
- **Transformation:** Performs data cleaning, validation, and schema formatting using pandas.
- **Loading:** Appends transformed data into a PostgreSQL relational database.

🛠️ Tech Stack
- **Language:** Python 3.x
- **Data Manipulation:** pandas
- **Database:** PostgreSQL
- **Version Control:** Git & GitHub

📁 Repository Structure
- `generar_datos.py`: Script to generate raw sample datasets for testing.
- `etl_script.py`: Core script containing the data cleaning and loading pipeline logic.
- `.gitignore`: Excludes environment variables, local configs, and temporary files.

⚙️ How to Run

1. **Clone the repository:**
   ```bash
   git clone [https://github.com/JuanP124/etl-pipeline-clientes.git](https://github.com/JuanP124/etl-pipeline-clientes.git)
   cd etl-pipeline-clientes
