import sys
import os
import glob
import json
import logging
import hashlib
import time
from datetime import datetime, timedelta
from typing import Callable, Optional

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from dotenv import load_dotenv

logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

load_dotenv()

TEMPLATE_WORD_PATH = r'inputs\template_report.docx'
OUTPUT_PATH = r'outputs'
CACHE_FILE = 'outputs/.ai_cache.json'

ProgressCallback = Optional[Callable[[float, str], None]]


class ValidationError(Exception):
    """Raised when the tickets Excel file has data problems that must be fixed before continuing."""
    def __init__(self, errors: list):
        self.errors = errors
        super().__init__("; ".join(errors))


def _report(on_progress: ProgressCallback, fraction: float, message: str):
    logging.info(message)
    if on_progress:
        on_progress(fraction, message)


def get_latest_excel() -> str:
    """Automatically finds the most recent excel file in the inputs folder."""
    excel_files = glob.glob(r'inputs\*.xlsx')
    excel_files = [f for f in excel_files if not os.path.basename(f).startswith('~$')]

    if not excel_files:
        logging.error("No valid Excel file was found in the 'inputs' folder.")
        sys.exit(1)

    latest_file = max(excel_files, key=os.path.getctime)
    logging.info(f"Excel file automatically detected: {latest_file}")
    return latest_file


def detect_errors(ex_df: pd.DataFrame) -> list:
    """Detects common errors in the tickets DataFrame. Returns a list of error messages (empty if none)."""
    errors = []

    if 'Ticket Id' in ex_df.columns:
        ticket_counts = ex_df['Ticket Id'].value_counts()
        duplicates = ticket_counts[ticket_counts > 1]
        if not duplicates.empty:
            for t_id, count in duplicates.items():
                errors.append(f"El ticket {t_id} está duplicado {count} veces")

    if 'Subject' in ex_df.columns and 'Group' in ex_df.columns:
        error2_mask = ex_df['Subject'].isna() | ex_df['Group'].isna()
        if error2_mask.any():
            if 'Ticket Id' in ex_df.columns:
                tickets_err2 = ex_df.loc[error2_mask, 'Ticket Id'].tolist()
                errors.append(f"Los siguientes tickets tienen Subject o Group en blanco: {tickets_err2}")
            else:
                errors.append(f"Hay {error2_mask.sum()} tickets con Subject o Group en blanco.")

    if 'Status' in ex_df.columns and 'Resolved Time' in ex_df.columns:
        error3_mask = ex_df['Status'].isin(['Closed', 'Resolved']) & ex_df['Resolved Time'].isna()
        if error3_mask.any():
            if 'Ticket Id' in ex_df.columns:
                tickets_err3 = ex_df.loc[error3_mask, 'Ticket Id'].tolist()
                errors.append(f"Los siguientes tickets están cerrados/resueltos pero sin 'Resolved Time': {tickets_err3}")
            else:
                errors.append(f"Hay {error3_mask.sum()} tickets cerrados/resueltos sin 'Resolved Time'.")

    return errors


def calculate_average_minutes(ex_df: pd.DataFrame, column: str, ticket_count: int) -> int:
    """Returns the average of a time column in whole minutes (0 if not computable)."""
    if column not in ex_df.columns or ticket_count == 0:
        return 0
    times = ex_df[column].dropna().astype(str)
    times_td = pd.to_timedelta(times, errors='coerce')
    total_time = times_td.sum()
    average = total_time / ticket_count
    return int(average.total_seconds()) // 60


def calculate_average_time(ex_df: pd.DataFrame, column: str, ticket_count: int) -> str:
    return f"{calculate_average_minutes(ex_df, column, ticket_count)} minutes"


def sla_rate(ex_df: pd.DataFrame, column: str, ticket_count: int) -> float:
    """Returns the percentage (0-100) of rows marked 'Within SLA' in the given column."""
    if column not in ex_df.columns or ticket_count == 0:
        return 0.0
    within_sla = (ex_df[column] == 'Within SLA').sum()
    return round((within_sla / ticket_count) * 100, 2)


def achieved_mark(actual: float, target: float, higher_is_better: bool = True) -> str:
    met = actual >= target if higher_is_better else actual <= target
    return "✅" if met else "❌"


def calc_percent(count: int, total: int) -> str:
    if total > 0:
        return f"{(count / total * 100):.2f}"
    return "0.00"


def count_item(ex_df: pd.DataFrame, item_name: str) -> int:
    if 'Item' in ex_df.columns:
        return (ex_df['Item'] == item_name).sum()
    return 0


def calculate_metrics(excel_df: pd.DataFrame, target_items: list) -> dict:
    ticket_count = len(excel_df)
    metrics = {
        'cantidad_tickets': ticket_count
    }

    if 'Agent interactions' in excel_df.columns:
        tickets_resolved_in_first_contact = excel_df['Agent interactions'].isin([1, 2]).sum()
    else:
        tickets_resolved_in_first_contact = 0
    metrics['tickets_resolved_in_first_contact'] = tickets_resolved_in_first_contact
    fr_contact_rate = float(calc_percent(tickets_resolved_in_first_contact, ticket_count))
    metrics['percentage_of_resolved_tickets_in_fr'] = f"{fr_contact_rate}%"

    if 'Survey Result' in excel_df.columns:
        surveys_answered = excel_df['Survey Result'].notna().sum()
    else:
        surveys_answered = 0
    metrics['surveys_answered'] = surveys_answered
    metrics['percentage_of_surveys_answered'] = f"{calc_percent(surveys_answered, ticket_count)}%"

    avg_fr_minutes = calculate_average_minutes(excel_df, 'First Response Time (in Hrs)', ticket_count)
    avg_resolution_minutes = calculate_average_minutes(excel_df, 'Resolution Time (in Hrs)', ticket_count)
    metrics['avg_fr_time'] = f"{avg_fr_minutes} minutes"
    metrics['avg_resolution'] = f"{avg_resolution_minutes} minutes"

    metrics['fr_sla_percent'] = sla_rate(excel_df, 'First Response Status', ticket_count)
    resolution_sla_percent = sla_rate(excel_df, 'Resolution Status', ticket_count)
    metrics['sla_resolution'] = f"{resolution_sla_percent}%"

    # Targets defined by the Experience team's service goals.
    metrics['fr_contact_achieved'] = achieved_mark(fr_contact_rate, 80)
    metrics['fr_time_achieved'] = achieved_mark(avg_fr_minutes, 700, higher_is_better=False)
    metrics['resolution_time_achieved'] = achieved_mark(avg_resolution_minutes, 1200, higher_is_better=False)
    metrics['sla_achieved'] = achieved_mark(resolution_sla_percent, 85)

    for item in target_items:
        key = item['key']
        item_name = item['name']
        count = count_item(excel_df, item_name)
        metrics[f"{key}_tickets"] = count
        metrics[f"percentage_of_{key}_tickets"] = calc_percent(count, ticket_count)

    if 'Item' in excel_df.columns:
        blank_tickets = excel_df['Item'].isna().sum()
    else:
        blank_tickets = 0
    metrics['blank_tickets'] = blank_tickets
    metrics['percentage_of_blank_tickets'] = calc_percent(blank_tickets, ticket_count)

    # Platform/tool breakdown table, driven entirely by config.json so the report
    # never drifts out of sync when a platform is added or removed from the config.
    target_items_display = [
        {
            'name': item['name'],
            'tickets': metrics[f"{item['key']}_tickets"],
            'percent': metrics[f"percentage_of_{item['key']}_tickets"],
        }
        for item in target_items
    ]
    target_items_display.append({
        'name': 'Not Provided',
        'tickets': blank_tickets,
        'percent': metrics['percentage_of_blank_tickets'],
    })
    # Only show platforms that actually received tickets this period, ordered by volume.
    target_items_display = sorted(
        (item for item in target_items_display if item['tickets'] > 0),
        key=lambda item: item['tickets'],
        reverse=True,
    )
    metrics['target_items_display'] = target_items_display

    if 'Status' in excel_df.columns:
        resolved_closed_tickets = excel_df['Status'].isin(['Resolved', 'Closed']).sum()
    else:
        resolved_closed_tickets = 0
    metrics['resolved_closed_tickets'] = resolved_closed_tickets
    metrics['percentage_of_resolved_tickets'] = calc_percent(resolved_closed_tickets, ticket_count)

    priority_levels = ['Low', 'Medium', 'High', 'Urgent']
    priorities_data = []

    if 'Priority' in excel_df.columns:
        priority_counts = excel_df['Priority'].str.capitalize().value_counts()
    else:
        priority_counts = pd.Series(dtype=int)

    for level in priority_levels:
        count = int(priority_counts.get(level.capitalize(), 0))
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

    if 'Nature' in excel_df.columns:
        nature_counts = excel_df['Nature'].value_counts().to_dict()
    else:
        nature_counts = {}
    metrics['nature_counts'] = nature_counts

    return metrics


def load_image(template: DocxTemplate, filename: str, width_mm: int = 150):
    path = os.path.join('inputs', filename)
    if os.path.exists(path):
        return InlineImage(template, path, width=Mm(width_mm))
    else:
        logging.warning(f"Image not found '{path}'")
        return ""


def load_cache() -> dict:
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_cache(cache_data: dict):
    os.makedirs(os.path.dirname(CACHE_FILE), exist_ok=True)
    with open(CACHE_FILE, 'w', encoding='utf-8') as f:
        json.dump(cache_data, f, ensure_ascii=False, indent=4)


def call_gemini_cached(prompt: str, default_text: str) -> str:
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

    max_retries = 3
    for attempt in range(max_retries):
        try:
            client = genai.Client(api_key=api_key.strip())
            resp = client.models.generate_content(
                model='gemini-2.5-flash',
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
                wait_time = float(match.group(1)) + 1.0 if match else 60.0
                logging.warning(f"Quota Limit 429. Retrying in {wait_time:.0f} seconds... (Attempt {attempt + 1}/{max_retries})")
                time.sleep(wait_time)
            else:
                logging.error(f"Error generating with AI: {e}")
                return default_text

    return default_text


def analyze_ai_patterns(df: pd.DataFrame) -> str:
    logging.info("Generating Key Common patterns...")
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        logging.warning("Gemini API key not found in .env file")
        return ""

    col_desc = 'Description' if 'Description' in df.columns else 'Subject'
    col_item = 'Item' if 'Item' in df.columns else None

    if col_desc and col_item:
        data = df[[col_item, col_desc]].dropna().head(100).to_dict('records')
    else:
        logging.warning("Columns required for AI analysis were not found.")
        return ""

    prompt = (
        "You are a data analyst. Analyze the following list of technical support tickets (each has an 'Item' and a 'Description'). "
        "Identify and summarize the key common patterns and the most frequent problems. "
        "Write the summary in English, ensuring it is clear and concise for a management report. "
        "IMPORTANT: Return only plain text, DO NOT use markdown formats like bold (**) or italics (*). Use dashes (-) for a single flat list of findings. "
        "Do NOT include any title, heading, or section label (e.g. do not write 'Key Common Patterns' or 'Most Frequent Problems') and do NOT group findings under sub-headings — "
        "start directly with the first bullet point.\n\n"
        f"Tickets: {json.dumps(data)}"
    )

    return call_gemini_cached(prompt, "")


def generate_ai_introduction(month_name: str, year: int, total_tickets: int) -> str:
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
    logging.info("Generating Survey (CSAT) summary...")

    if 'Survey Result' not in df.columns:
        logging.warning("No 'Survey Result' column found.")
        return "No survey data available for this period."

    survey_map = {'2/2': 'Awesome', '1/2': 'Satisfactory', '0/2': 'Not good'}
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


BRAND_BLUE = '#1560BD'
BRAND_PURPLE = '#7B2D8E'
BRAND_QUALITATIVE = ['#1560BD', '#7B2D8E', '#5B8DEF', '#B36BC2', '#9CA3AF']


def generate_graphs(df: pd.DataFrame):
    """Generates the charts automatically using matplotlib and seaborn. Returns the list of filenames written."""
    logging.info("Generating graphs automatically...")
    os.makedirs('inputs', exist_ok=True)
    sns.set_theme(style="whitegrid")
    written = []

    if 'Status' in df.columns:
        plt.figure(figsize=(6, 6))
        status_counts = df['Status'].value_counts()
        if not status_counts.empty:
            colors = BRAND_QUALITATIVE[:len(status_counts)]
            plt.pie(
                status_counts, labels=status_counts.index, autopct='%1.1f%%',
                startangle=140, colors=colors, wedgeprops={'linewidth': 0},
            )
            plt.title('Tickets by Status', color='#1F2A6B', fontweight='bold')
        plt.tight_layout()
        plt.savefig(r'inputs\tickets_by_status.png')
        plt.close()
        written.append('tickets_by_status.png')

    if 'Nature' in df.columns:
        plt.figure(figsize=(8, 5))
        nature_counts = df['Nature'].value_counts().head(5)
        if not nature_counts.empty:
            palette = list(reversed(sns.light_palette(BRAND_BLUE, n_colors=len(nature_counts) + 1)[1:]))
            sns.barplot(x=nature_counts.values, y=nature_counts.index, hue=nature_counts.index, palette=palette, legend=False)
            plt.title('Top 5 Request Categories (Nature)', color='#1F2A6B', fontweight='bold')
            plt.xlabel('Number of Tickets')
            plt.ylabel('')
        plt.tight_layout()
        plt.savefig(r'inputs\tickets_by_nature.png')
        plt.close()
        written.append('tickets_by_nature.png')

    if 'Created Time' in df.columns:
        plt.figure(figsize=(10, 5))
        try:
            df_temp = df.copy()
            df_temp['Created Time'] = pd.to_datetime(df_temp['Created Time'])
            date_counts = df_temp['Created Time'].dt.date.value_counts().sort_index()
            if not date_counts.empty:
                sns.lineplot(x=date_counts.index, y=date_counts.values, marker="o", color=BRAND_BLUE)
                plt.title('Tickets Received by Date', color='#1F2A6B', fontweight='bold')
                plt.xlabel('Date')
                plt.ylabel('Tickets')
                plt.xticks(rotation=45)
            plt.tight_layout()
            plt.savefig(r'inputs\tickets_by_date.png')
            written.append('tickets_by_date.png')
        except Exception as e:
            logging.warning(f"Could not generate date graph: {e}")
        finally:
            plt.close()

    if 'Priority' in df.columns and 'First Response Time (in Hrs)' in df.columns:
        plt.figure(figsize=(8, 5))
        try:
            df_temp = df.copy()
            df_temp['First_Res_Num'] = pd.to_timedelta(df_temp['First Response Time (in Hrs)'].astype(str), errors='coerce').dt.total_seconds() / 3600
            avg_resp = df_temp.groupby('Priority')['First_Res_Num'].mean().dropna().sort_values()
            if not avg_resp.empty:
                palette = list(reversed(sns.light_palette(BRAND_PURPLE, n_colors=len(avg_resp) + 1)[1:]))
                sns.barplot(x=avg_resp.index, y=avg_resp.values, hue=avg_resp.index, palette=palette, legend=False)
                plt.title('Average First Response Time by Priority (Hours)', color='#1F2A6B', fontweight='bold')
                plt.xlabel('Priority')
                plt.ylabel('Hours')
            plt.tight_layout()
            plt.savefig(r'inputs\tickets_by_response_time.png')
            written.append('tickets_by_response_time.png')
        except Exception as e:
            logging.warning(f"Could not generate response time graph: {e}")
        finally:
            plt.close()

    if 'Resolved Time' in df.columns:
        plt.figure(figsize=(10, 5))
        try:
            df_temp = df.copy()
            df_temp['Resolved Time'] = pd.to_datetime(df_temp['Resolved Time'])
            resolved_counts = df_temp['Resolved Time'].dt.date.value_counts().sort_index()
            if not resolved_counts.empty:
                sns.lineplot(x=resolved_counts.index, y=resolved_counts.values, marker="s", color=BRAND_PURPLE)
                plt.title('Tickets Resolved by Date', color='#1F2A6B', fontweight='bold')
                plt.xlabel('Date')
                plt.ylabel('Tickets Resolved')
                plt.xticks(rotation=45)
            plt.tight_layout()
            plt.savefig(r'inputs\tickets_by_closed.png')
            written.append('tickets_by_closed.png')
        except Exception as e:
            logging.warning(f"Could not generate closed tickets graph: {e}")
        finally:
            plt.close()

    if 'Survey Result' in df.columns:
        plt.figure(figsize=(6, 6))
        try:
            survey_map = {'2/2': 'Awesome', '1/2': 'Satisfactory', '0/2': 'Not good'}
            mapped_survey = df['Survey Result'].dropna().astype(str).map(survey_map).dropna()
            survey_counts = mapped_survey.value_counts()
            if not survey_counts.empty:
                colors = BRAND_QUALITATIVE[:len(survey_counts)]
                plt.pie(
                    survey_counts, labels=survey_counts.index, autopct='%1.1f%%',
                    startangle=140, colors=colors, wedgeprops={'linewidth': 0},
                )
                plt.title('Customer Satisfaction Survey Results', color='#1F2A6B', fontweight='bold')
            plt.tight_layout()
            plt.savefig(r'inputs\tickets_by_survey.png')
            written.append('tickets_by_survey.png')
        except Exception as e:
            logging.warning(f"Could not generate survey graph: {e}")
        finally:
            plt.close()

    def _sla_donut(column: str, title: str, color: str, filename: str):
        if column not in df.columns:
            return
        ticket_count = len(df)
        percent = sla_rate(df, column, ticket_count)
        plt.figure(figsize=(4, 4))
        try:
            remainder = max(0.0, 100 - percent)
            plt.pie(
                [percent, remainder],
                colors=[color, '#E5E7EB'],
                startangle=90,
                counterclock=False,
                wedgeprops={'width': 0.35, 'linewidth': 0},
            )
            plt.text(0, 0, f"{percent:.0f} %", ha='center', va='center', fontsize=22, fontweight='bold')
            plt.title(title, color='#1F2A6B', fontsize=15, fontweight='bold')
            plt.legend(['Within SLA'], loc='lower center', bbox_to_anchor=(0.5, -0.05), frameon=False)
            plt.tight_layout()
            plt.savefig(os.path.join('inputs', filename))
            written.append(filename)
        except Exception as e:
            logging.warning(f"Could not generate {filename}: {e}")
        finally:
            plt.close()

    _sla_donut('First Response Status', 'First Response SLA', BRAND_BLUE, 'fr_sla.png')
    _sla_donut('Resolution Status', 'Resolution SLA', BRAND_PURPLE, 'r_sla.png')

    return written


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


def load_target_items(config_path: str = 'config.json') -> list:
    if os.path.exists(config_path):
        with open(config_path, 'r', encoding='utf-8') as f:
            config_data = json.load(f)
            return config_data.get('target_items', [])
    logging.warning("config.json not found, platform metrics will be empty.")
    return []


def generate_report(
    tickets_file: str,
    template_path: str = TEMPLATE_WORD_PATH,
    output_dir: str = OUTPUT_PATH,
    on_progress: ProgressCallback = None,
) -> dict:
    """
    Runs the full pipeline: validate -> graphs -> metrics -> AI text -> render Word.
    Returns a dict with 'output_path', 'metrics' and 'month_label'.
    Raises ValidationError if the Excel data has blocking issues.
    """
    _report(on_progress, 0.05, "Leyendo el archivo Excel...")
    excel_df = pd.read_excel(tickets_file, sheet_name='Sheet_0')

    _report(on_progress, 0.12, "Validando los datos...")
    errors = detect_errors(excel_df)
    if errors:
        raise ValidationError(errors)

    prev_month, prev_month_name, prev_month_year = get_report_date()

    _report(on_progress, 0.25, "Generando las gráficas...")
    generate_graphs(excel_df)

    _report(on_progress, 0.35, "Calculando métricas...")
    target_items = load_target_items()
    metrics = calculate_metrics(excel_df, target_items)

    _report(on_progress, 0.45, "Analizando patrones con IA...")
    ai_patterns = analyze_ai_patterns(excel_df)
    time.sleep(1)

    _report(on_progress, 0.55, "Generando introducción con IA...")
    intro_text = generate_ai_introduction(prev_month_name, prev_month_year, metrics['cantidad_tickets'])
    time.sleep(1)

    _report(on_progress, 0.65, "Generando resumen de prioridades...")
    priority_summary = generate_ai_priority_summary(prev_month_name, prev_month_year, metrics['cantidad_tickets'], metrics['priorities_data'])
    time.sleep(1)

    _report(on_progress, 0.75, "Generando resumen de categorías (Nature)...")
    nature_summary = generate_ai_nature_summary(prev_month_name, prev_month_year, metrics['cantidad_tickets'], metrics['nature_counts'])
    time.sleep(1)

    _report(on_progress, 0.85, "Generando resumen de encuestas (CSAT)...")
    survey_summary = generate_ai_survey_summary(prev_month_name, prev_month_year, metrics['cantidad_tickets'], excel_df)

    _report(on_progress, 0.92, "Renderizando el documento Word...")
    docx_template = DocxTemplate(template_path)

    img_tickets_by_status = load_image(docx_template, 'tickets_by_status.png', width_mm=90)
    img_tickets_by_nature = load_image(docx_template, 'tickets_by_nature.png', width_mm=150)
    img_tickets_by_date = load_image(docx_template, 'tickets_by_date.png', width_mm=160)
    img_tickets_by_response_time = load_image(docx_template, 'tickets_by_response_time.png', width_mm=160)
    img_tickets_by_closed = load_image(docx_template, 'tickets_by_closed.png', width_mm=160)
    img_tickets_by_survey = load_image(docx_template, 'tickets_by_survey.png', width_mm=90)
    img_fr_sla = load_image(docx_template, 'fr_sla.png', width_mm=75)
    img_r_sla = load_image(docx_template, 'r_sla.png', width_mm=75)

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
        'fr_contact_achieved': metrics['fr_contact_achieved'],
        'fr_time_achieved': metrics['fr_time_achieved'],
        'resolution_time_achieved': metrics['resolution_time_achieved'],
        'sla_achieved': metrics['sla_achieved'],
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
        'fr_sla': img_fr_sla,
        'r_sla': img_r_sla,
        'ai_patterns': ai_patterns,
        'priorities_data': metrics['priorities_data'],
        'target_items_display': metrics['target_items_display'],
    }

    for item in target_items:
        key = item['key']
        context[f"{key}_tickets"] = str(metrics.get(f"{key}_tickets", 0))
        context[f"percentage_of_{key}_tickets"] = metrics.get(f"percentage_of_{key}_tickets", "0.00")

    docx_template.render(context)

    os.makedirs(output_dir, exist_ok=True)
    output_filename = f"Report_tickets_{prev_month:02d}_{prev_month_year}.docx"
    output_filepath = os.path.join(output_dir, output_filename)
    docx_template.save(output_filepath)

    _report(on_progress, 1.0, "¡Reporte generado exitosamente!")

    return {
        'output_path': output_filepath,
        'metrics': metrics,
        'month_label': f"{prev_month_name} {prev_month_year}",
    }
