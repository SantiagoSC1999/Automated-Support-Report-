# Excel to Word Report Automation 📊📄

An automated Python tool designed to convert exported ticketing data (from Freshservice or similar helpdesks) into a professional, AI-powered Word Document report.

This script parses an `.xlsx` file, calculates operational metrics, generates beautiful charts automatically using `matplotlib` and `seaborn`, and uses Google's **Gemini AI** to write analytical summaries (Key Patterns, Introductions, Priorities) before injecting everything seamlessly into a Word template.

## Features ✨

- **Data Processing:** Automatically detects the latest Excel file in the `inputs/` folder, cleans data, and detects anomalies (missing fields, duplicate tickets).
- **Automated Graphics:** Generates 5 fully-styled charts based on your data (Pie, Bar, Line charts) without any manual intervention.
- **AI-Powered Summaries:** Uses Google Gemini (via prompt engineering) to analyze data and write contextual paragraphs for your report.
- **Smart Caching:** Local caching mechanism to save your AI API quota during iterative testing.
- **Dynamic Platform Tracking:** Easily configure which software platforms or tools you want to track by editing `config.json` without touching the Python code.

## Prerequisites 🛠️

1. **Python 3.10+** installed on your system.
2. A free [Google Gemini API Key](https://aistudio.google.com/app/apikey).

## Installation 🚀

1. Clone this repository:
   ```bash
   git clone https://github.com/yourusername/excel-to-word-report.git
   cd excel-to-word-report
   ```

2. Create a virtual environment and install the required dependencies:
   ```bash
   python -m venv .venv
   # On Windows:
   .\.venv\Scripts\activate
   # On Mac/Linux:
   source .venv/bin/activate

   pip install -r requirements.txt
   ```

3. Create a `.env` file in the root directory and add your Gemini API Key:
   ```env
   GEMINI_API_KEY=your_api_key_here
   ```

## Configuration ⚙️

### 1. `config.json`
Define the names of the tools, products, or platforms you want to track in the `config.json` file. The tool will calculate the metrics for these specific items.

```json
{
  "target_items": [
    {
      "key": "platform_one",
      "name": "Actual Name in Excel"
    }
  ]
}
```
*Note: The `key` will be the prefix for your Word template variables (e.g., `{{platform_one_tickets}}` and `{{percentage_of_platform_one_tickets}}`).*

### 2. The Word Template
Place your Word template inside the `inputs/` folder and name it `template_report.docx`.
Use Jinja2 syntax (e.g. `{{ variable_name }}`) wherever you want the script to inject data, AI text, or charts.

### 3. The Excel Data
Export your tickets to an `.xlsx` format and drop it in the `inputs/` folder. The script will automatically pick up the most recent Excel file it finds.
*Required Columns: `Status`, `Nature`, `Priority`, `Item`, `Created Time`, `Resolved Time`, `First Response Time (in Hrs)`.*

## Usage 💻

Run the script from your terminal:

```bash
python fromexceltoword.py
```

Check the `outputs/` folder! Your new report `Report_tickets_MM_YYYY.docx` will be waiting for you.

## License 📜
MIT License
