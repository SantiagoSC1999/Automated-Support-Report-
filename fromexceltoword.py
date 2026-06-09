import sys
import os
import glob
import json
import logging
import hashlib
import time
from datetime import datetime, timedelta

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from dotenv import load_dotenv

# Basic logging configuration
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

load_dotenv() # Load variables from .env file

TEMPLATE_WORD_PATH = r'inputs\template_report.docx'
OUTPUT_PATH = r'outputs'

def get_latest_excel() -> str:
    """Automatically finds the most recent excel file in the inputs folder."""
    excel_files = glob.glob(r'inputs\*.xlsx')
    
    # Ignore hidden temporary files that Excel creates when opening a file (e.g., ~$file.xlsx)
    excel_files = [f for f in excel_files if not os.path.basename(f).startswith('~$')]
    
    if not excel_files:
        logging.error("No valid Excel file was found in the 'inputs' folder.")
        sys.exit(1)

    # Keep the most recently created file
    latest_file = max(excel_files, key=os.path.getctime)
    logging.info(f"Excel file automatically detected: {latest_file}")
    return latest_file

def detect_errors(ex_df: pd.DataFrame):
    """Detects common errors in the tickets DataFrame and aborts if necessary."""
    err_1, err_2, err_3 = False, False, False

    # Error 1: Duplicated tickets (the same Ticket ID should not appear twice)
    if 'Ticket Id' in ex_df.columns:
        ticket_counts = ex_df['Ticket Id'].value_counts()
        duplicates = ticket_counts[ticket_counts > 1]
        if not duplicates.empty:
            for t_id, count in duplicates.items():
                logging.error(f"Ticket {t_id} is duplicated {count} times!!")
            err_1 = True

    # Error 2: Tickets without Subject or Group assigned
    if 'Subject' in ex_df.columns and 'Group' in ex_df.columns:
        error2_mask = ex_df['Subject'].isna() | ex_df['Group'].isna()
        if error2_mask.any():
            if 'Ticket Id' in ex_df.columns:
                tickets_err2 = ex_df.loc[error2_mask, 'Ticket Id'].tolist()
                logging.error(f"The following tickets have a blank Subject or Group: {tickets_err2}")
            else:
                logging.error(f"There are {error2_mask.sum()} tickets with a blank Subject or Group.")
            err_2 = True

    # Error 3: Tickets marked as resolved/closed but without a resolution date
    if 'Status' in ex_df.columns and 'Resolved Time' in ex_df.columns:
        error3_mask = ex_df['Status'].isin(['Closed', 'Resolved']) & ex_df['Resolved Time'].isna()
        if error3_mask.any():
            if 'Ticket Id' in ex_df.columns:
                tickets_err3 = ex_df.loc[error3_mask, 'Ticket Id'].tolist()
                logging.error(f"The following tickets are closed/resolved but have no 'Resolved Time': {tickets_err3}")
            else:
                logging.error(f"There are {error3_mask.sum()} closed/resolved tickets with no 'Resolved Time'.")
            err_3 = True

    # If any error is detected, halt program execution
    if err_1 or err_2 or err_3:
        logging.error("PLEASE CORRECT ERRORS IN THE EXCEL FILE TO CONTINUE EXECUTION...")
        sys.exit(1)
    else:
        logging.info("No errors detected in the tickets :)")

def calculate_average_time(ex_df: pd.DataFrame, column: str, ticket_count: int) -> str:
    """Calculates the average time for a given column in 'X minutes' format."""
    if column not in ex_df.columns:
        return "0 minutes"
    times = ex_df[column].dropna().astype(str)
    times_td = pd.to_timedelta(times, errors='coerce')
    total_time = times_td.sum()
    if ticket_count > 0:
        average = total_time / ticket_count
        total_seconds = int(average.total_seconds())
        total_minutes = total_seconds // 60
        return f"{total_minutes} minutes"
    return "0 minutes"

def calc_percent(count: int, total: int) -> str:
    """Calculates the percentage (without the % symbol) formatted to two decimal places."""
    if total > 0:
        return f"{(count / total * 100):.2f}"
    return "0.00"

def count_item(ex_df: pd.DataFrame, item_name: str) -> int:
    """Counts the occurrence of a specific item."""
    if 'Item' in ex_df.columns:
        return (ex_df['Item'] == item_name).sum()
    return 0

def calculate_metrics(excel_df: pd.DataFrame, target_items: list) -> dict:
    """Calculates all metrics from the tickets DataFrame."""
    ticket_count = len(excel_df)
    metrics = {
        'cantidad_tickets': ticket_count  # Kept var name for context injection compatibility
    }

    # 1. Tickets resolved in the first contact (interactions 1 or 2)
    if 'Agent interactions' in excel_df.columns:
        tickets_resolved_in_first_contact = excel_df['Agent interactions'].isin([1, 2]).sum()
    else:
        tickets_resolved_in_first_contact = 0
    metrics['tickets_resolved_in_first_contact'] = tickets_resolved_in_first_contact
    metrics['percentage_of_resolved_tickets_in_fr'] = f"{calc_percent(tickets_resolved_in_first_contact, ticket_count)}%"

    # 1b. Answered surveys and their percentage
    if 'Survey Result' in excel_df.columns:
        surveys_answered = excel_df['Survey Result'].notna().sum()
    else:
        surveys_answered = 0
    metrics['surveys_answered'] = surveys_answered
    metrics['percentage_of_surveys_answered'] = f"{calc_percent(surveys_answered, ticket_count)}%"

    # 2 & 3. Time averages
    metrics['avg_fr_time'] = calculate_average_time(excel_df, 'First Response Time (in Hrs)', ticket_count)
    metrics['avg_resolution'] = calculate_average_time(excel_df, 'Resolution Time (in Hrs)', ticket_count)

    # 4. SLA Resolution %
    if 'Resolution Status' in excel_df.columns:
        sla_count = (excel_df['Resolution Status'] == 'Within SLA').sum()
        metrics['sla_resolution'] = f"{calc_percent(sla_count, ticket_count)}%"
    else:
        metrics['sla_resolution'] = "0.00%"

    # 5. Tickets by Item defined in config.json
    for item in target_items:
        key = item['key']
        item_name = item['name']
        count = count_item(excel_df, item_name)
        metrics[f"{key}_tickets"] = count
        metrics[f"percentage_of_{key}_tickets"] = calc_percent(count, ticket_count)

    # Blank tickets
    if 'Item' in excel_df.columns:
        blank_tickets = excel_df['Item'].isna().sum()
    else:
        blank_tickets = 0
    metrics['blank_tickets'] = blank_tickets
    metrics['percentage_of_blank_tickets'] = calc_percent(blank_tickets, ticket_count)

    # 6. Resolved/Closed Tickets
    if 'Status' in excel_df.columns:
        resolved_closed_tickets = excel_df['Status'].isin(['Resolved', 'Closed']).sum()
    else:
        resolved_closed_tickets = 0
    metrics['resolved_closed_tickets'] = resolved_closed_tickets
    metrics['percentage_of_resolved_tickets'] = calc_percent(resolved_closed_tickets, ticket_count)

    # 7. Data for the Priorities Table (Priority)
    priority_levels = ['Low', 'Medium', 'High', 'Urgent']
    priorities_data = []
    
    if 'Priority' in excel_df.columns:
        # Normalize to title case to avoid issues like "low" vs "Low"
        priority_counts = excel_df['Priority'].str.capitalize().value_counts()
    else:
        priority_counts = pd.Series(dtype=int)

    for level in priority_levels:
        count = int(priority_counts.get(level.capitalize(), 0))
        # In the sample table there are no decimals, so we round to integer
        if ticket_count > 0:
            percent_str = f"{int(round((count / ticket_count) * 100))}%"
        else:
            percent_str = "0%"

        priorities_data.append({
            'level': level,
            'count': count,
            'percent': percent_str
        })
    
    metrics['priorities_data'] = priorities_data

    # 8. Data for Nature
    if 'Nature' in excel_df.columns:
        # Count and convert to dictionary (already ordered from highest to lowest by default in pandas)
        nature_counts = excel_df['Nature'].value_counts().to_dict()
    else:
        nature_counts = {}
    metrics['nature_counts'] = nature_counts

    return metrics

def load_image(template: DocxTemplate, filename: str, width_mm: int = 150):
    """Loads an image into the Word template, returns an empty string if it does not exist."""
    path = os.path.join('inputs', filename)
    if os.path.exists(path):
        return InlineImage(template, path, width=Mm(width_mm))
    else:
        logging.warning(f"Image not found '{path}'")
        return ""

CACHE_FILE = 'outputs/.ai_cache.json'

def load_cache() -> dict:
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return {}
    return {}

def save_cache(cache_data: dict):
    os.makedirs(os.path.dirname(CACHE_FILE), exist_ok=True)
    with open(CACHE_FILE, 'w', encoding='utf-8') as f:
        json.dump(cache_data, f, ensure_ascii=False, indent=4)

def call_gemini_cached(prompt: str, default_text: str) -> str:
    """Makes the request to Gemini, using local cache to save quota and automatic retries."""
    cache = load_cache()
    prompt_hash = hashlib.md5(prompt.encode('utf-8')).hexdigest()
    
    if prompt_hash in cache:
        logging.info("Using Gemini response from local cache (Quota saved!).")
        return cache[prompt_hash]

    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        logging.warning("Gemini API key not found.")
        return default_text

    from google import genai
    import time
    
    max_retries = 3
    for attempt in range(max_retries):
        try:
            client = genai.Client(api_key=api_key.strip())
            resp = client.models.generate_content(
                model='gemini-3.5-flash',
                contents=prompt
            )
            response_text = resp.text.strip()
            cache[prompt_hash] = response_text
            save_cache(cache)
            logging.info("AI response generated successfully.")
            return response_text
        except Exception as e:
            if "503" in str(e) and attempt < max_retries - 1:
                logging.warning(f"API 503 Error. Retrying in 3 seconds... (Attempt {attempt + 1}/{max_retries})")
                time.sleep(3)
            elif "429" in str(e) and attempt < max_retries - 1:
                import re
                match = re.search(r'retry in ([\d\.]+)s', str(e))
                if match:
                    wait_time = float(match.group(1)) + 1.0 # Add 1 second extra to be safe
                else:
                    wait_time = 60.0 # Default to 60s if unable to parse
                logging.warning(f"Quota Limit 429. Retrying in {wait_time:.0f} seconds... (Attempt {attempt + 1}/{max_retries})")
                time.sleep(wait_time)
            else:
                logging.error(f"Error generating with AI: {e}")
                return default_text
    
    return default_text

def analyze_ai_patterns(df: pd.DataFrame) -> str:
    """Generates a pattern analysis using the Gemini API."""
    logging.info("Generating Key Common patterns...")
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        logging.warning("Gemini API key not found in .env file")
        return "" # We return empty so as not to damage the report

    col_desc = 'Description' if 'Description' in df.columns else 'Subject'
    col_item = 'Item' if 'Item' in df.columns else None

    if col_desc and col_item:
        data = df[[col_item, col_desc]].dropna().head(100).to_dict('records')
    else:
        logging.warning("Columns required for AI analysis were not found.")
        return ""

    prompt = (
        "You are a data analyst. Analyze the following list of technical support tickets (each has an 'Item' and a 'Description'). "
        "Identify and summarize the 'Key Common Patterns' and the most frequent problems. "
        "Write the summary in English, ensuring it is clear and concise for a management report. "
        "IMPORTANT: Return only plain text, DO NOT use markdown formats like bold (**) or italics (*). Use dashes (-) for lists.\n\n"
        f"Tickets: {json.dumps(data)}"
    )

    return call_gemini_cached(prompt, "")

def generate_ai_introduction(month_name: str, year: int, total_tickets: int) -> str:
    """Generates a dynamic introduction using the Gemini API."""
    logging.info("Generating introduction...")
    
    default_text = (
        f"As part of the Experience team's ongoing commitment to enhancing our applications, we continuously monitor requests for improvements and modifications. "
        f"With a strong focus on reducing errors and optimizing the user experience, we perform detailed analyses of key operational metrics. "
        f"This includes assessing the volume of support requests received through our ticketing system (Freshservice) in {month_name} {year}, where we managed a total of {total_tickets} tickets. "
        f"In this monthly report, we provide a comprehensive overview of the activities managed by the Experience area. The report outlines the impact of key indicators, including ticket volumes, resolution metrics, and lessons learned. Through this analysis, we aim to reinforce our commitment to operational excellence and continuous improvement."
    )

    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        logging.warning("Gemini API key not found, default introduction will be used.")
        return default_text

    prompt = (
        f"Act as a technical manager. Write a brief introduction paragraph (maximum 15 lines) for a monthly support report. "
        f"The report is for the 'Experience team' and covers the month of {month_name} {year}. "
        f"It must mention that the team continuously monitors requests for improvements and modifications to enhance applications, focusing on reducing errors and optimizing user experience. "
        f"It must mention that this report provides an overview of the {total_tickets} support requests received through our ticketing system (Freshservice). "
        f"It must outline that the report covers key indicators like ticket volumes and lessons learned to reinforce commitment to operational excellence. "
        f"The tone should be professional and formal. DO NOT mention JIRA. ONLY mention Freshservice. "
        f"Make the phrasing slightly unique and dynamic so it doesn't read identically every single month, but keep the core message intact. "
        f"IMPORTANT: Return only plain text, NO markdown formats like bold (**) or italics (*)."
    )

    return call_gemini_cached(prompt, default_text)

def generate_ai_priority_summary(month_name: str, year: int, total_tickets: int, priorities_data: list) -> str:
    """Generates a dynamic summary of the priorities table using the Gemini API."""
    logging.info("Generating priorities summary...")
    
    counts = {p['level']: p['count'] for p in priorities_data}
    low = counts.get('Low', 0)
    medium = counts.get('Medium', 0)
    high = counts.get('High', 0)
    urgent = counts.get('Urgent', 0)

    default_text = (
        f"Across the chosen period in {month_name} {year}, a significant number of tickets were managed with a variety of priority levels. "
        f"Out of the total {total_tickets} processed tickets, the majority ({low}) were designated as low priority, "
        f"while medium accounted for {medium} tickets, high for {high} tickets, and urgent tickets accounted for {urgent}."
    )

    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        logging.warning("Gemini API key not found, default priority summary will be used.")
        return default_text

    prompt = (
        f"Act as a data analyst. Write a brief summary paragraph (maximum 5 lines) analyzing the priority levels of support tickets for {month_name} {year}. "
        f"The total number of tickets is {total_tickets}. Here is the breakdown: "
        f"Low: {low}, Medium: {medium}, High: {high}, Urgent: {urgent}. "
        f"Mention the chosen period ({month_name} {year}), the total tickets, and summarize the distribution of priorities naturally in English. "
        f"Make it read professionally for a report. Vary the wording slightly so it's not identical every month. "
        f"IMPORTANT: Return only plain text, NO markdown formats like bold (**) or italics (*)."
    )

    return call_gemini_cached(prompt, default_text)

def generate_ai_nature_summary(month_name: str, year: int, total_tickets: int, nature_counts: dict) -> str:
    """Generates a dynamic summary of the Nature breakdown using the Gemini API."""
    logging.info("Generating Nature summary (Request categories breakdown)...")
    
    default_text = (
        f"This data indicates the volume of requests broken down by their nature. "
        f"In {month_name} {year}, out of {total_tickets} total tickets, "
        f"we managed requests across various categories to support our active platforms."
    )

    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        logging.warning("Gemini API key not found, default Nature summary will be used.")
        return default_text

    # Filter only the top 5 so as not to saturate the prompt and give the AI only the most important data
    top_nature = dict(list(nature_counts.items())[:5]) if nature_counts else {}

    prompt = (
        f"Act as a data analyst. Write a brief summary paragraph (maximum 4 lines) analyzing the 'Request categories breakdown' (Nature of tickets) for {month_name} {year}. "
        f"The total number of tickets is {total_tickets}. Here is the breakdown of the top categories by volume: "
        f"{json.dumps(top_nature)}. "
        f"Start the paragraph with a sentence similar to 'This data indicates that the highest volume of requests were related with...' "
        f"Mention the top 1 or 2 categories and their proportion out of the total {total_tickets} tickets. "
        f"Make it read professionally for a report. Vary the wording slightly so it's not identical every month. "
        f"IMPORTANT: Return only plain text, NO markdown formats like bold (**) or italics (*)."
    )

    return call_gemini_cached(prompt, default_text)

def generate_ai_survey_summary(month_name: str, year: int, total_tickets: int, df: pd.DataFrame) -> str:
    """Generates a dynamic summary of Customer Satisfaction using the Gemini API."""
    logging.info("Generating Survey (CSAT) summary...")
    
    if 'Survey Result' not in df.columns:
        logging.warning("No 'Survey Result' column found.")
        return "No survey data available for this period."

    survey_map = {'2/2': 'Awesome', '1/2': 'Satisfactory', '0/2': 'Not good'}
    # Replace valid values and ignore nan or undefined
    mapped_survey = df['Survey Result'].dropna().astype(str).map(survey_map).dropna()
    survey_counts = mapped_survey.value_counts().to_dict()
    
    awesome = survey_counts.get('Awesome', 0)
    satisfactory = survey_counts.get('Satisfactory', 0)
    not_good = survey_counts.get('Not good', 0)
    answered = awesome + satisfactory + not_good

    default_text = (
        f"In {month_name} {year}, out of {total_tickets} tickets, we received {answered} survey responses. "
        f"The feedback indicated {awesome} 'Awesome' ratings, {satisfactory} 'Satisfactory' ratings, and {not_good} 'Not good' ratings."
    )

    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        logging.warning("Gemini API key not found, default survey summary will be used.")
        return default_text

    prompt = (
        f"Act as a customer success manager. Write a brief summary paragraph (maximum 4 lines) analyzing the customer satisfaction survey results for {month_name} {year}. "
        f"Total tickets: {total_tickets}. Total surveys answered: {answered}. "
        f"Breakdown of ratings: Awesome: {awesome}, Satisfactory: {satisfactory}, Not good: {not_good}. "
        f"Analyze these results professionally. If 'Awesome' is high, highlight the team's excellent service. If 'Not good' is present, mention a commitment to improving. "
        f"Make it read professionally for a report. Vary the wording slightly so it's not identical every month. "
        f"IMPORTANT: Return only plain text, NO markdown formats like bold (**) or italics (*)."
    )

    return call_gemini_cached(prompt, default_text)

def generate_graphs(df: pd.DataFrame):
    """Generates the 5 graphs automatically using matplotlib and seaborn."""
    logging.info("Generating graphs automatically...")
    os.makedirs('inputs', exist_ok=True)
    sns.set_theme(style="whitegrid")
    
    # 1. tickets_by_status.png
    if 'Status' in df.columns:
        plt.figure(figsize=(6, 6))
        status_counts = df['Status'].value_counts()
        if not status_counts.empty:
            plt.pie(status_counts, labels=status_counts.index, autopct='%1.1f%%', startangle=140, colors=sns.color_palette("pastel"))
            plt.title('Tickets by Status')
        plt.tight_layout()
        plt.savefig(r'inputs\tickets_by_status.png')
        plt.close()

    # 2. tickets_by_nature.png
    if 'Nature' in df.columns:
        plt.figure(figsize=(8, 5))
        nature_counts = df['Nature'].value_counts().head(5)
        if not nature_counts.empty:
            sns.barplot(x=nature_counts.values, y=nature_counts.index, hue=nature_counts.index, palette="viridis", legend=False)
            plt.title('Top 5 Request Categories (Nature)')
            plt.xlabel('Number of Tickets')
            plt.ylabel('')
        plt.tight_layout()
        plt.savefig(r'inputs\tickets_by_nature.png')
        plt.close()

    # 3. tickets_by_date.png
    if 'Created Time' in df.columns:
        plt.figure(figsize=(10, 5))
        try:
            df_temp = df.copy()
            df_temp['Created Time'] = pd.to_datetime(df_temp['Created Time'])
            date_counts = df_temp['Created Time'].dt.date.value_counts().sort_index()
            if not date_counts.empty:
                sns.lineplot(x=date_counts.index, y=date_counts.values, marker="o", color="blue")
                plt.title('Tickets Received by Date')
                plt.xlabel('Date')
                plt.ylabel('Tickets')
                plt.xticks(rotation=45)
            plt.tight_layout()
            plt.savefig(r'inputs\tickets_by_date.png')
        except Exception as e:
            logging.warning(f"Could not generate date graph: {e}")
        finally:
            plt.close()

    # 4. tickets_by_response_time.png
    if 'Priority' in df.columns and 'First Response Time (in Hrs)' in df.columns:
        plt.figure(figsize=(8, 5))
        try:
            df_temp = df.copy()
            df_temp['First_Res_Num'] = pd.to_timedelta(df_temp['First Response Time (in Hrs)'].astype(str), errors='coerce').dt.total_seconds() / 3600
            avg_resp = df_temp.groupby('Priority')['First_Res_Num'].mean().dropna().sort_values()
            if not avg_resp.empty:
                sns.barplot(x=avg_resp.index, y=avg_resp.values, hue=avg_resp.index, palette="magma", legend=False)
                plt.title('Average First Response Time by Priority (Hours)')
                plt.xlabel('Priority')
                plt.ylabel('Hours')
            plt.tight_layout()
            plt.savefig(r'inputs\tickets_by_response_time.png')
        except Exception as e:
            logging.warning(f"Could not generate response time graph: {e}")
        finally:
            plt.close()

    # 5. tickets_by_closed.png
    if 'Resolved Time' in df.columns:
        plt.figure(figsize=(10, 5))
        try:
            df_temp = df.copy()
            df_temp['Resolved Time'] = pd.to_datetime(df_temp['Resolved Time'])
            resolved_counts = df_temp['Resolved Time'].dt.date.value_counts().sort_index()
            if not resolved_counts.empty:
                sns.lineplot(x=resolved_counts.index, y=resolved_counts.values, marker="s", color="green")
                plt.title('Tickets Resolved by Date')
                plt.xlabel('Date')
                plt.ylabel('Tickets Resolved')
                plt.xticks(rotation=45)
            plt.tight_layout()
            plt.savefig(r'inputs\tickets_by_closed.png')
        except Exception as e:
            logging.warning(f"Could not generate closed tickets graph: {e}")
        finally:
            plt.close()

    # 6. tickets_by_survey.png
    if 'Survey Result' in df.columns:
        plt.figure(figsize=(6, 6))
        try:
            survey_map = {'2/2': 'Awesome', '1/2': 'Satisfactory', '0/2': 'Not good'}
            mapped_survey = df['Survey Result'].dropna().astype(str).map(survey_map).dropna()
            survey_counts = mapped_survey.value_counts()
            if not survey_counts.empty:
                # Use a pie chart for surveys
                plt.pie(survey_counts, labels=survey_counts.index, autopct='%1.1f%%', startangle=140, colors=sns.color_palette("pastel"))
                plt.title('Customer Satisfaction Survey Results')
            plt.tight_layout()
            plt.savefig(r'inputs\tickets_by_survey.png')
        except Exception as e:
            logging.warning(f"Could not generate survey graph: {e}")
        finally:
            plt.close()

def get_report_date():
    """Calculates the previous month and the corresponding year."""
    today = datetime.now()
    first_day_current_month = today.replace(day=1)
    last_day_previous_month = first_day_current_month - timedelta(days=1)
    
    prev_month = last_day_previous_month.month
    months_en = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
    prev_month_name = months_en[prev_month - 1]
    prev_month_year = last_day_previous_month.year

    return prev_month, prev_month_name, prev_month_year

def main():
    # 1. Search Excel
    tickets_file = get_latest_excel()
    
    # 2. Load data
    excel_df = pd.read_excel(tickets_file, sheet_name='Sheet_0')

    # 3. Validate data
    detect_errors(excel_df)
    logging.info("Validation completed. Starting Word processing...")

    # 4. Get dates
    prev_month, prev_month_name, prev_month_year = get_report_date()

    # 5. Generate Graphs Automatically
    generate_graphs(excel_df)

    # 5b. Load Items configuration
    config_path = 'config.json'
    if os.path.exists(config_path):
        with open(config_path, 'r', encoding='utf-8') as f:
            config_data = json.load(f)
            target_items = config_data.get('target_items', [])
    else:
        # Fallback if config.json does not exist
        target_items = []
        logging.warning("config.json not found, platform metrics will be empty.")

    # 6. Calculate Metrics
    metrics = calculate_metrics(excel_df, target_items)
    
    # 7. Execute AI calls (Sequential and Modular)
    logging.info("Launching requests to Gemini sequentially and safely...")
    
    ai_patterns = analyze_ai_patterns(excel_df)
    time.sleep(2) # Small pause to avoid saturating Google's free tier
    
    intro_text = generate_ai_introduction(prev_month_name, prev_month_year, metrics['cantidad_tickets'])
    time.sleep(2)
    
    priority_summary = generate_ai_priority_summary(prev_month_name, prev_month_year, metrics['cantidad_tickets'], metrics['priorities_data'])
    time.sleep(2)
    
    nature_summary = generate_ai_nature_summary(prev_month_name, prev_month_year, metrics['cantidad_tickets'], metrics['nature_counts'])
    time.sleep(2)
    
    survey_summary = generate_ai_survey_summary(prev_month_name, prev_month_year, metrics['cantidad_tickets'], excel_df)

    # 8. Load Word template and Images
    docx_template = DocxTemplate(TEMPLATE_WORD_PATH)
    
    img_tickets_by_status = load_image(docx_template, 'tickets_by_status.png', width_mm=80)
    img_tickets_by_nature = load_image(docx_template, 'tickets_by_nature.png', width_mm=80)
    img_tickets_by_date = load_image(docx_template, 'tickets_by_date.png', width_mm=150)
    img_tickets_by_response_time = load_image(docx_template, 'tickets_by_response_time.png', width_mm=150)
    img_tickets_by_closed = load_image(docx_template, 'tickets_by_closed.png', width_mm=150)
    img_tickets_by_survey = load_image(docx_template, 'tickets_by_survey.png', width_mm=100)

    # 9. Build context for the template
    context = {
        'introduction': intro_text,
        'priority_summary': priority_summary,
        'nature_summary': nature_summary,
        'survey_summary': survey_summary,
        'month': prev_month_name,
        'year': str(prev_month_year),
        'tickets_prev_month': str(metrics['cantidad_tickets']),
        'tickets_total': str(metrics['cantidad_tickets']),
        'tickets_resolved_in_first_contact': str(metrics['tickets_resolved_in_first_contact']),
        'percentage_of_resolved_tickets_in_fr': metrics['percentage_of_resolved_tickets_in_fr'],
        'surveys_answered': str(metrics['surveys_answered']),
        'percentage_of_surveys_answered': metrics['percentage_of_surveys_answered'],
        'avg_fr_time': metrics['avg_fr_time'],
        'avg_resolution': metrics['avg_resolution'],
        'sla_resolution': metrics['sla_resolution'],
        'blank_tickets': str(metrics['blank_tickets']),
        'percentage_of_blank_tickets': metrics['percentage_of_blank_tickets'],
        'resolved_closed_tickets': str(metrics['resolved_closed_tickets']),
        'percentage_of_resolved_tickets': metrics['percentage_of_resolved_tickets'],
        'tickets_by_status': img_tickets_by_status,
        'tickets_by_nature': img_tickets_by_nature,
        'tickets_by_date': img_tickets_by_date,
        'tickets_by_response_time': img_tickets_by_response_time,
        'tickets_by_closed': img_tickets_by_closed,
        'tickets_by_survey': img_tickets_by_survey,
        'ai_patterns': ai_patterns,
        'priorities_data': metrics['priorities_data'],
    }

    # Add dynamic metrics for platforms configured in config.json
    for item in target_items:
        key = item['key']
        context[f"{key}_tickets"] = str(metrics.get(f"{key}_tickets", 0))
        context[f"percentage_of_{key}_tickets"] = metrics.get(f"percentage_of_{key}_tickets", "0.00")

    # 10. Render and Save Report
    docx_template.render(context)
    
    # Ensure output folder exists
    os.makedirs(OUTPUT_PATH, exist_ok=True)
    
    output_filename = f"Report_tickets_{prev_month:02d}_{prev_month_year}.docx"
    output_filepath = os.path.join(OUTPUT_PATH, output_filename)
    
    docx_template.save(output_filepath)
    logging.info(f"Report successfully generated at: {output_filepath}")

if __name__ == "__main__":
    main()