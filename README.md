from email import encoders
from email.mime.base import MIMEBase
from DBconnection import db_connections
import pandas as pd
from datetime import datetime, time
import os
import sys
import time as time_module  # For sleep function
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import logging
from logging.handlers import RotatingFileHandler
import boto3  # For AWS SES email and S3
import pytz  # For timezone handling
import io

# ============================================================================
# LOGGING CONFIGURATION
# ============================================================================
USE_S3 = True

# Captured ONCE at startup — ensures setup_logging() and upload_log_to_s3() use the same timestamp
# Explicitly use IST for logging and file naming so they make sense to the team
IST_TZ = pytz.timezone('Asia/Kolkata')
RUN_TIMESTAMP = datetime.now(IST_TZ).strftime('%Y-%m-%dT%H-%M-%S')  # e.g. 2026-03-05T12-30-00 (IST)
LOG_DIR = "/tmp/logs" if (USE_S3 and os.name != 'nt') else "logs"

# Custom formatter to force IST in all log lines
class ISTFormatter(logging.Formatter):
    def converter(self, timestamp):
        return datetime.fromtimestamp(timestamp, IST_TZ).timetuple()

def setup_logging():
    """
    Configure logging with CONSOLE handler only.
    No main validation log file is created at startup.
    Per-job log files (e.g. PS_PatientConnect_2026-03-09_11-08-00.log) are created
    only when a schedule time matches via add_job_log_handler().
    """
    # Create logger
    logger = logging.getLogger('ScheduledValidation')
    logger.setLevel(logging.INFO)

    # Clear existing handlers
    logger.handlers = []

    # Console handler only — no file handler at startup
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_formatter = ISTFormatter('%(asctime)s - %(levelname)s - %(message)s')
    console_handler.setFormatter(console_formatter)
    logger.addHandler(console_handler)

    return logger


def upload_log_to_s3():
    """
    Disabled: Main validation log file is no longer created.
    Per-job logs are uploaded inside remove_job_log_handler() instead.
    """
    pass  # No-op: per-job logs handle their own S3 upload


def add_job_log_handler(job_name, stage_label, scheduled_time=""):
    """
    Add a job-specific file handler to the global logger.
    Uses a filename like: {job_name}_{date}_{scheduled_time}.log
    If running in S3 mode, it attempts to download the existing log to append to it.
    """
    if not os.path.exists(LOG_DIR):
        os.makedirs(LOG_DIR)

    # Sanitize inputs
    safe_name = job_name.replace(' ', '_').replace('/', '_').replace('\\', '_')
    safe_time = str(scheduled_time).replace(':', '-').replace(' ', '_')
    run_date = datetime.now(IST_TZ).strftime('%Y-%m-%d')
    
    # Filename includes scheduled time so Stage 2 can find Stage 1's log for the same specific run
    log_filename = f"{safe_name}_{run_date}"
    if safe_time:
        log_filename += f"_{safe_time}"
    log_filename += ".log"
    
    log_filepath = os.path.join(LOG_DIR, log_filename)

    # If USE_S3 is True, try to download existing log from S3 to append
    if USE_S3:
        try:
            s3_client = boto3.client("s3", region_name=EMAIL_CONFIG["aws_region"])
            s3_key = f"logs/{log_filename}"
            # Check if file exists in S3
            s3_client.head_object(Bucket=S3_CONFIG["bucket"], Key=s3_key)
            # If it exists, download it
            s3_client.download_file(S3_CONFIG["bucket"], s3_key, log_filepath)
            logger.info(f"[OK] Downloaded existing job log from S3 to append: {log_filename}")
        except Exception:
            # File doesn't exist or other error, starts fresh
            pass

    # Open in append mode 'a' so we can combine stages if they run in the same process or from S3
    handler = logging.FileHandler(log_filepath, mode='a')
    handler.setLevel(logging.INFO)
    handler.setFormatter(ISTFormatter('%(asctime)s - %(levelname)s - %(message)s'))

    logger.addHandler(handler)
    logger.info(f"\n{'='*40}\n[LOG] STAGE: {stage_label} | Job: {job_name}\n{'='*40}")
    
    return handler, log_filepath


def remove_job_log_handler(handler, log_filepath):
    """
    Remove a job-specific file handler from the logger and upload its log to S3.
    """
    log_filename = os.path.basename(log_filepath)
    logger.info(f"[LOG] Job log stage finished -> {log_filename}")
    logger.removeHandler(handler)
    handler.close()

    # Upload per-job log to S3
    if USE_S3:
        try:
            s3      = boto3.client("s3", region_name=EMAIL_CONFIG["aws_region"])
            s3_key  = f"logs/{log_filename}"     # -> s3://bucket/logs/<job>_<date>.log
            s3.upload_file(log_filepath, S3_CONFIG["bucket"], s3_key)
            logger.info(f"[OK] Job log synced to S3 -> s3://{S3_CONFIG['bucket']}/{s3_key}")
        except Exception as e:
            logger.warning(f"Could not upload job log to S3: {e}")


# Initialize logger
logger = setup_logging()

# ============================================================================
# CONFIGURATION
# ============================================================================

# --------------------------------------------------------------------------
# ★ CHANGE HERE ★  S3 Configuration (used when USE_S3 = True)
# --------------qq  ------------------------------------------------------------

# ★ CHANGE 1: Set to True when deploying on AWS Glue / EC2 / Lambda
#             Leave False when running locally


S3_CONFIG = {
    # ★ CHANGE 2: Replace with your actual S3 bucket name
    "bucket": "eda-qaresults-eversana",

    "prefix": "Replicate_Config",
}

# --------------------------------------------------------------------------
# Local File paths (used when USE_S3 = False  i.e. running locally)
# No changes needed here unless you move your config folder
# --------------------------------------------------------------------------
CONFIG_DIR = "Replicate_Config"
SCHEDULE_CONFIG_FILE  = os.path.join(CONFIG_DIR, "Replicate_schedule_Config_test.xlsx")

# ★ This filename MUST match the exact filename in S3 under 'Replicate_Config/' prefix
INPUT_EXCEL_FILENAME  = "Replicate_Tables.xlsx"   # ← S3 key: Replicate_Config/Replicate_Tables.xlsx
INPUT_EXCEL           = os.path.join(CONFIG_DIR, INPUT_EXCEL_FILENAME)

# STAGE1_EXCEL, STAGE2_EXCEL, OUTPUT_EXCEL, FAILED_EXCEL removed - paths are now dynamic per use case.


# BATCH_SIZE removed - processing all tables per use case at once for simplicity

# --------------------------------------------------------------------------
# ★ CHANGE HERE ★  Email Configuration - AWS SES Only
# --------------------------------------------------------------------------

EMAIL_CONFIG = {
    # ★ CHANGE 4: AWS region where SES is configured
    "aws_region": "us-east-2",
    # ★ CHANGE 5: The FROM address — must be verified in AWS SES console
    "ses_sender": "EDA-QA Team<noreply@eversana.com>",
}

# Time tolerance for schedule matching (in minutes)
# Increased to 1 to ensure triggers don't miss if the script starts a few seconds late.
TIME_TOLERANCE_MINUTES = 1


# ============================================================================
# S3 HELPER FUNCTIONS
# ============================================================================

def _s3_key(filename):
    """Build full S3 key from filename using the configured prefix."""
    return f"{S3_CONFIG['prefix']}/{filename}"



def read_excel_from_s3(filename, silent=False):
    """
    Read an Excel file from S3 and return as DataFrame.

    Args:
        filename: Just the filename, e.g. 'Replicate_schedule_Config_test.xlsx'
        silent: If True, don't log an error if file is not found (useful for optional files)

    Returns:
        pd.DataFrame or None on error
    """
    
    try:
        s3 = boto3.client("s3", region_name=EMAIL_CONFIG["aws_region"])
        key = _s3_key(filename)
        logger.debug(f"Reading s3://{S3_CONFIG['bucket']}/{key}")
        obj = s3.get_object(Bucket=S3_CONFIG["bucket"], Key=key)
        return pd.read_excel(io.BytesIO(obj["Body"].read()))
    except Exception as e:
        if not silent:
            logger.error(f"Failed to read s3://{S3_CONFIG['bucket']}/{_s3_key(filename)}: {e}")
        return None


def write_excel_to_s3(df, filename):
    """
    Write a DataFrame as an Excel file to S3.

    Args:
        df:       DataFrame to write
        filename: Just the filename, e.g. 'count_validation_results.xlsx'

    Returns:
        bool: True if successful
    """
    import io
    try:
        s3 = boto3.client("s3", region_name=EMAIL_CONFIG["aws_region"])
        key = _s3_key(filename)
        buffer = io.BytesIO()
        df.to_excel(buffer, index=False)
        buffer.seek(0)
        s3.put_object(Bucket=S3_CONFIG["bucket"], Key=key, Body=buffer.getvalue())
        logger.debug(f"Written s3://{S3_CONFIG['bucket']}/{key}")
        return True
    except Exception as e:
        logger.error(f"Failed to write s3://{S3_CONFIG['bucket']}/{_s3_key(filename)}: {e}")
        return False


def delete_s3_file(filename):
    """
    Delete a file from S3.

    Args:
        filename: Just the filename, e.g. 'count_validation_results.xlsx'

    Returns:
        bool: True if successful
    """
    try:
        s3 = boto3.client("s3", region_name=EMAIL_CONFIG["aws_region"])
        key = _s3_key(filename)
        s3.delete_object(Bucket=S3_CONFIG["bucket"], Key=key)
        logger.debug(f"Deleted s3://{S3_CONFIG['bucket']}/{key}")
        return True
    except Exception as e:
        logger.error(f"Failed to delete s3://{S3_CONFIG['bucket']}/{_s3_key(filename)}: {e}")
        return False


# ============================================================================
# UTILITY FUNCTIONS
# ============================================================================

def load_pipeline_schedules():
    """
    Load pipeline schedules from Excel file (local or S3 depending on USE_S3 flag)

    Returns:
        list: List of schedule dictionaries with keys:
              replicated_jobs, job_1_time_ist, job_2_time_ist, email_recipients (optional)
    """
    try:
        if USE_S3:
            # ── Read from S3 ──
            logger.debug(f"Loading schedule config from S3: {S3_CONFIG['bucket']}/{_s3_key('Replicate_schedule_Config_test.xlsx')}")
            df = read_excel_from_s3("Replicate_schedule_Config_test.xlsx")
            if df is None:
                logger.error("Could not load schedule config from S3. Check bucket/prefix/file.")
                return []
        else:
            # ── Read from local file ──
            if not os.path.exists(SCHEDULE_CONFIG_FILE):
                logger.error(f"Configuration file not found: {SCHEDULE_CONFIG_FILE}")
                logger.info(f"Run: python {sys.argv[0]} create-config")
                return []
            df = pd.read_excel(SCHEDULE_CONFIG_FILE)

        logger.debug(f"Loaded {len(df)} schedules")

        # Validate required columns (Support both IST and UTC naming, but we will treat as UTC)
        col_map = {
            'job_1': 'job_1_time_utc' if 'job_1_time_utc' in df.columns else 'job_1_time_ist',
            'job_2': 'job_2_time_utc' if 'job_2_time_utc' in df.columns else 'job_2_time_ist'
        }
        
        required_columns = ['replicated_jobs', col_map['job_1'], col_map['job_2']]
        missing_columns = [col for col in required_columns if col not in df.columns]

        if missing_columns:
            logger.error(f"Missing columns in schedule config: {missing_columns}")
            return []
            
        # Standardize keys to _utc for internal use
        schedules = []
        for _, row in df.iterrows():
            item = {
                'replicated_jobs': row['replicated_jobs'],
                'job_1_time_utc': row[col_map['job_1']],
                'job_2_time_utc': row[col_map['job_2']],
                'email_recipients': row.get('email_recipients', '')
            }
            schedules.append(item)
        return schedules

    except Exception as e:
        logger.error(f"Error loading pipeline schedules: {e}", exc_info=True)
        return []


def load_email_recipients():
    """
    Load the global email recipient list from the 'email_recipients' sheet
    inside Replicate_schedule_Config_test.xlsx.

    Sheet structure (one row per recipient):
        | email                    |
        |--------------------------|  
        | alice@eversana.com       |
        | bob@eversana.com         |
        | manager@eversana.com     |

    Works with both local files and S3 (respects the USE_S3 flag).

    Returns:
        list[str]: List of email addresses.  Empty list if sheet is missing
                   or no valid addresses are found.
    """
    try:
        if USE_S3:
            # For S3 we read the whole workbook via BytesIO
            import io as _io
            s3 = boto3.client("s3", region_name=EMAIL_CONFIG["aws_region"])
            key = _s3_key(os.path.basename(SCHEDULE_CONFIG_FILE))
            obj = s3.get_object(Bucket=S3_CONFIG["bucket"], Key=key)
            raw = _io.BytesIO(obj["Body"].read())
            xl = pd.ExcelFile(raw)
        else:
            if not os.path.exists(SCHEDULE_CONFIG_FILE):
                logger.error(f"Schedule config not found: {SCHEDULE_CONFIG_FILE}")
                return []
            xl = pd.ExcelFile(SCHEDULE_CONFIG_FILE)

        # Check the sheet exists
        if "email_recipients" not in xl.sheet_names:
            logger.error(
                "Sheet 'email_recipients' not found in config file. "
                "Please add a sheet named 'email_recipients' with a column 'email'."
            )
            return []

        df = xl.parse("email_recipients")

        if "email" not in df.columns:
            logger.error(
                "Column 'email' not found in 'email_recipients' sheet. "
                "The sheet must have a column named 'email'."
            )
            return []

        # Clean up: drop blanks, strip whitespace
        recipients = [
            str(e).strip()
            for e in df["email"].dropna()
            if str(e).strip() != ""
        ]

        if not recipients:
            logger.error("'email_recipients' sheet exists but contains no email addresses.")
            return []

        logger.debug(f"Loaded {len(recipients)} email recipient(s): {recipients}")
        return recipients

    except Exception as e:
        logger.error(f"Failed to load email recipients: {e}", exc_info=True)
        return []


def parse_time_string(time_str):
    """
    Parse time string in HH:MM or HH:MM:SS format to datetime.time object
    
    Args:
        time_str: Time string in "HH:MM" or "HH:MM:SS" format
    
    Returns:
        datetime.time object or None if error
    """
    try:
        # Convert to string and split by ':'
        time_parts = str(time_str).split(':')
        
        # Handle both HH:MM and HH:MM:SS formats
        if len(time_parts) == 2:
            # HH:MM format
            hour, minute = map(int, time_parts)
            return time(hour=hour, minute=minute)
        elif len(time_parts) == 3:
            # HH:MM:SS format - ignore seconds
            hour, minute, _ = map(int, time_parts)
            return time(hour=hour, minute=minute)
        else:
            logger.error(f"Invalid time format '{time_str}': Expected HH:MM or HH:MM:SS")
            return None
    except Exception as e:
        logger.error(f"Error parsing time '{time_str}': {e}")
        return None


def check_schedule_match(current_time, scheduled_time_str, tolerance_minutes=TIME_TOLERANCE_MINUTES):
    """
    Check if current time matches scheduled time within tolerance
    
    Args:
        current_time: datetime object with current time
        scheduled_time_str: Scheduled time as string "HH:MM"
        tolerance_minutes: Tolerance in minutes
    
    Returns:
        tuple: (bool, int) - (True if matches, difference in minutes)
    """
    scheduled_time = parse_time_string(scheduled_time_str)
    if not scheduled_time:
        return False, 9999
    
    # Convert current time to minutes since midnight
    current_minutes = current_time.hour * 60 + current_time.minute
    scheduled_minutes = scheduled_time.hour * 60 + scheduled_time.minute
    
    # Check if within tolerance
    diff = abs(current_minutes - scheduled_minutes)
    return diff <= tolerance_minutes, diff


def filter_data_by_usecase(df, use_case):
    """
    Filter DataFrame to include only rows for specified use case.
    Does NOT remove rows with blank table_names (they will be dynamically expanded).
    
    Args:
        df: Input DataFrame
        use_case: Use case name to filter
    
    Returns:
        Filtered DataFrame
    """
    # Fix the missing 'table_name' column header gracefully
    if 'table_name' not in df.columns:
        if 'Unnamed: 3' in df.columns.astype(str):
            df.rename(columns={'Unnamed: 3': 'table_name'}, inplace=True)
        else:
            # If the column got saved as index 3 due to no header:
            if len(df.columns) > 3:
                df.rename(columns={df.columns[3]: 'table_name'}, inplace=True)
            else:
                df['table_name'] = pd.NA
            
    # Forward fill the replicated_jobs column
    if 'replicated_jobs' in df.columns:
        df['replicated_jobs'] = df['replicated_jobs'].ffill()
    else:
        # Failsafe if the column isn't found
        df['replicated_jobs'] = df.iloc[:, 0].ffill()
    
    # Filter by use case
    df_filtered = df[df['replicated_jobs'] == use_case].copy()
    
    return df_filtered


def fetch_schema_tables_info(conn, db, schema):
    """
    Unified query function that fetches all table data for a given schema at once.
    Returns: list of dicts with 'TABLE_NAME', 'ROW_COUNT', 'TABLE_TYPE'
    """
    try:
        query = f"""
        SELECT TABLE_NAME, ROW_COUNT, TABLE_TYPE 
        FROM "{db}".INFORMATION_SCHEMA.TABLES 
        WHERE TABLE_SCHEMA = '{schema.upper()}'
        """
        logger.info(f"[QUERY] Fetching metadata for schema '{schema}' from database '{db}'...")
        cursor = conn.cursor()
        cursor.execute(query)
        results = cursor.fetchall()
        logger.info(f"[OK] Retrieved {len(results)} table(s) metadata from {db}.{schema}")
        cursor.close()
        
        tables_info = []
        for row in results:
            tables_info.append({
                'TABLE_NAME': str(row[0]).upper() if row[0] else '',
                'ROW_COUNT': int(row[1]) if row[1] is not None else 0,
                'TABLE_TYPE': str(row[2]).upper() if row[2] else ''
            })
        return tables_info
    except Exception as e:
        logger.error(f"Error fetching schema info for {db}.{schema}: {e}")
        return []

def expand_dynamic_tables(conn, df_filtered):
    """
    Expands rows with missing table_name by getting all 'BASE TABLE's from the schema.
    """
    expanded_rows = []
    
    for _, row in df_filtered.iterrows():
        # Check if table_name is provided and not empty
        if pd.notna(row['table_name']) and str(row['table_name']).strip():
            # Already explicit
            expanded_rows.append(row.to_dict())
        else:
            # Dynamic: fetch all tables for this db and schema using the unified query
            db = row['database_name']
            schema = row['schema_name']
            
            tables_info = fetch_schema_tables_info(conn, db, schema)
            
            for t_info in tables_info:
                if t_info['TABLE_TYPE'] == 'BASE TABLE':
                    # Create a new row for each core table
                    new_row = row.to_dict()
                    new_row['table_name'] = t_info['TABLE_NAME']
                    expanded_rows.append(new_row)
                    
    if expanded_rows:
        return pd.DataFrame(expanded_rows)
    else:
        return pd.DataFrame(columns=df_filtered.columns)

# ============================================================================
# DATABASE FUNCTIONS
# ============================================================================

def get_table_count(conn, database, schema, table):
    """
    Fetch row count for a specified database table using the unified query.
    """
    try:
        tables_info = fetch_schema_tables_info(conn, database, schema)
        for t_info in tables_info:
            if t_info['TABLE_NAME'] == table.upper():
                return t_info['ROW_COUNT']
        return 0
    except Exception as e:
        logger.error(f"Error fetching info schema count for {database}.{schema}.{table}: {e}")
        return None

def get_batch_table_counts(conn, tables_list):
    """
    Fetch row counts for multiple tables using the unified query.
    """
    try:
        if not tables_list:
            return {}
        
        # Group tables by unique (database, schema) pairs
        db_schemas = set((db, schema) for db, schema, _ in tables_list)
        
        counts_dict_upper = {}
        
        # Query info schema once per unique schema using the unified lookup
        for db, schema in db_schemas:
            logger.info(f"[ACTION] Batch processing metadata for schema: {db}.{schema}")
            tables_info = fetch_schema_tables_info(conn, db, schema)
            
            # Store returned info schema data using uppercase keys
            for t_info in tables_info:
                key = (db.upper(), schema.upper(), t_info['TABLE_NAME'])
                counts_dict_upper[key] = t_info['ROW_COUNT']
        
        # Create final dictionary matching the case of the input tables_list
        final_counts = {}
        for db, schema, table in tables_list:
            key_upper = (db.upper(), schema.upper(), table.upper())
            if key_upper in counts_dict_upper:
                final_counts[(db, schema, table)] = counts_dict_upper[key_upper]
        
        return final_counts
        
    except Exception as e:
        logger.warning(f"Batch info schema processing failed: {e}. Falling back to individual queries")
        return None


# ============================================================================
# EMAIL FUNCTIONS
# ============================================================================



def send_validation_notification(use_case, actual_count, expected_count, failed_df, recipients, status='FAILED'):
    """
    Send an email notification for a use case validation.
    Sends on both PASS and FAIL conditions using AWS SES.
    """
    try:
        if not recipients:
            logger.error("No email recipients found.")
            return False

        # Status icon and color
        status_icon = "\U0001f534" if status == 'FAILED' else "\U00002705"
        status_text = "FAILED" if status == 'FAILED' else "PASSED"
        title_color = "#d9534f" if status == 'FAILED' else "#5cb85c"

        est_timezone = pytz.timezone('US/Eastern')
        est_time     = datetime.now(est_timezone).strftime('%H:%M:%S')
        now_str      = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        msg = MIMEMultipart()
        msg["Subject"] = f"{status_icon} PROD_{status_text}_{use_case}"

        html_parts = ["<html><body>"]
        if status == 'FAILED':
            html_parts.append(f"<h3 style='color:{title_color};'>[WARNING] Count Validation Failed</h3>")
        else:
            html_parts.append(f"<h3 style='color:{title_color};'>[SUCCESS] Count Validation Passed</h3>")
        
        html_parts.append("<h4>Summary:</h4>")
        html_parts.append(f"<p><strong>Replicate Jobs:</strong> {use_case}</p>")
        html_parts.append(f"<p><strong>Total Tables Validated:</strong> {expected_count}</p>")
        
        if status == 'FAILED':
            html_parts.append(f"<p><strong>Failed Validations:</strong> {expected_count - actual_count}</p>")
        else:
            html_parts.append(f"<p><strong>Passed Validations:</strong> {actual_count}</p>")
            
        html_parts.append(f"<p><strong>Validation Date:</strong> {now_str}</p>")

        if status == 'FAILED' and failed_df is not None and len(failed_df) > 0:
            html_parts.append("<br><h4 style='color:#d9534f;'>Failed Tables Details:</h4>")
            html_parts.append(
                "<p>The following tables have counts that did not meet the validation criteria</p>"
            )

            detail = pd.DataFrame({
                "Replicate_Jobs":   failed_df['replicated_jobs'],
                "Table Name":       (failed_df['Database_Name'] + '.'
                                     + failed_df['Schema_Name']  + '.'
                                     + failed_df['Table_Name']),
                "Before_Count":  failed_df['First_Count'].apply(
                                        lambda x: f"{int(x):,}" if pd.notna(x) else 'N/A'),
                "After_Count":      failed_df['Second_Count'].apply(
                                        lambda x: f"{int(x):,}" if pd.notna(x) else 'N/A'),
                "Difference":       failed_df.apply(
                                        lambda r: (
                                            f"-{int(r['First_Count']) - int(r['Second_Count']):,}"
                                            if pd.notna(r['First_Count']) and pd.notna(r['Second_Count'])
                                            else 'N/A'
                                        ), axis=1),
                "Remarks":          failed_df['Remarks']
            })

            raw_html = detail.to_html(index=False, classes='table', border=1)
            styled   = (
                raw_html
                .replace('<table',
                         '<table style="border-collapse:collapse;width:100%;'
                         'font-family:Arial,sans-serif;font-size:14px;color:#000;"')
                .replace('<th>',
                         '<th style="background-color:#000;color:#fff;padding:10px;'
                         'text-align:left;border:2px solid #333;font-size:14px;">')
                .replace('<td>',
                         '<td style="padding:10px;text-align:left;'
                         'border:1px solid #aaa;color:#000;font-size:13px;">')
            )
            html_parts.append(styled)

        html_parts.append("</body></html>")
        msg.attach(MIMEText("".join(html_parts), "html"))

        ses_client = boto3.client("ses", region_name=EMAIL_CONFIG['aws_region'])
        ses_client.send_raw_email(
            Source=EMAIL_CONFIG['ses_sender'], 
            Destinations=recipients, 
            RawMessage={"Data": msg.as_bytes()}
        )
        
        logger.info(f"[OK] {status_text} email sent to {recipients}")
        return True
    except Exception as e:
        logger.error(f"Failed to send email via AWS SES: {e}", exc_info=True)
        return False


# ============================================================================
# CLEANUP FUNCTION - DELETE RESULTS AFTER ALL JOBS COMPLETE
# ============================================================================

def check_and_cleanup_completed_jobs():
    """
    Check if all scheduled jobs have completed their validation (Stages 2, 3, 4)
    and delete the count_validation_results.xlsx file if all are done.
    
    This prepares the system for the next day's run.
    
    Returns:
        bool: True if cleanup was performed, False otherwise
    """
    try:
        # This function is now legacy as Stage 4 handles cleanup per use case.
        return True
        
        # Load schedules from config
        schedules = load_pipeline_schedules()
        
        if not schedules:
            logger.debug("No schedules loaded. Cannot determine if cleanup is needed.")
            return False
        
        # Get all use cases from schedule config
        scheduled_use_cases = set([schedule.get('replicated_jobs', '') for schedule in schedules if schedule.get('replicated_jobs', '')])
        # This function is now fully redundant because Stage 4 cleans up per use case.
        # Returning True safely to avoid any FileNotFoundError for obsolete global files.
        return True
    except Exception as e:
        logger.error(f"Error during legacy cleanup check: {e}", exc_info=True)
        return False


# ============================================================================
# STAGE 1: CAPTURE FIRST COUNT
# ============================================================================

def capture_first_count_for_usecase(use_case):
    """
    STAGE 1: Capture first count for tables in specified use case
    
    Args:
        use_case: Use case name to process
    
    Returns:
        bool: True if successful, False otherwise
    """
    logger.info("="*80)
    logger.info(f"STAGE 1: CAPTURING FIRST COUNT - {use_case}")
    logger.info("="*80)
    
    try:
        # Read input Excel file
        if USE_S3:
            df_input = read_excel_from_s3(INPUT_EXCEL_FILENAME)
            if df_input is None:
                logger.error(f"Could not load {INPUT_EXCEL_FILENAME} from S3 "
                             f"(expected at s3://{S3_CONFIG['bucket']}/{S3_CONFIG['prefix']}/{INPUT_EXCEL_FILENAME}). "
                             f"Please upload the file to S3 first.")
                return False
        else:
            if not os.path.exists(INPUT_EXCEL):
                logger.error(f"Input file not found: {INPUT_EXCEL}")
                return False
            df_input = pd.read_excel(INPUT_EXCEL)
        
        # Filter by use case
        df_filtered = filter_data_by_usecase(df_input, use_case)
        
        # Connect to database
        logger.info(f"[ACTION] Connecting to Snowflake to fetch counts for {use_case}...")
        conn = db_connections("SNOWFLAKE", "Service")
        
        if not conn:
            logger.error("[ERROR] Database connection failed")
            return False
            
        logger.info("[OK] Database connection successful")
        
        # Expand dynamic tables (where table_name is left blank in excel)
        logger.info(f"[ACTION] Checking for dynamic tables in schemas for use case: {use_case}")
        df_filtered = expand_dynamic_tables(conn, df_filtered)
        
        if len(df_filtered) == 0:
            logger.warning(f"No tables found for use case: {use_case}")
            return False
        
        logger.info(f"Processing {len(df_filtered)} tables for '{use_case}'")
        logger.debug(f"Tables: {df_filtered['table_name'].tolist()}")
        
        # Prepare results list
        results = []
        
        # Note: We overwrite results for THIS use case, so no need to load existing
        # when paths are dynamic per use case.
        stage1_filename = f"{use_case}_Stage1.xlsx"
        stage1_filepath = os.path.join(CONFIG_DIR, stage1_filename)
        
        total_tables = len(df_filtered)
        logger.info(f"[ACTION] Processing all {total_tables} tables for metadata counts...")
        
        # Prepare list for counts lookup
        tables_to_query = []
        for _, row in df_filtered.iterrows():
            tables_to_query.append((row['database_name'], row['schema_name'], row['table_name']))
        
        # Highly efficient lookup (queries each schema only once)
        all_counts = get_batch_table_counts(conn, tables_to_query)
        
        # Process results
        for index, row in df_filtered.iterrows():
            db, sc, tb = row['database_name'], row['schema_name'], row['table_name']
            count = all_counts.get((db, sc, tb)) if all_counts else None
            
            # Fallback to individual if batch failed
            if count is None:
                count = get_table_count(conn, db, sc, tb)
                
            if count is not None:
                results.append({
                    'replicated_jobs': row['replicated_jobs'],
                    'Database_Name': db,
                    'Schema_Name': sc,
                    'Table_Name': tb,
                    'First_Count': count,
                    'Second_Count': None,
                    'Status': 'PENDING',
                    'First_Count_Timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    'Second_Count_Timestamp': None,
                    'Remarks': f'First count captured for use case: {use_case}'
                })
            else:
                results.append({
                    'replicated_jobs': row['replicated_jobs'],
                    'Database_Name': db,
                    'Schema_Name': sc,
                    'Table_Name': tb,
                    'First_Count': None,
                    'Second_Count': None,
                    'Status': 'ERROR',
                    'First_Count_Timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    'Second_Count_Timestamp': None,
                    'Remarks': 'Error fetching first count'
                })
        
        # Close connection
        conn.close()
        
        # Since paths are dynamic per use case, we just save the new results
        df_final = pd.DataFrame(results)
        
        # Save results to Excel (local or S3)
        stage1_filename = f"{use_case}_Stage1.xlsx"
        if USE_S3:
            write_excel_to_s3(df_final, stage1_filename)
            logger.info(f"  Stage 1 Results saved to: s3://{S3_CONFIG['bucket']}/{_s3_key(stage1_filename)}")
        else:
            stage1_filepath = os.path.join(CONFIG_DIR, stage1_filename)
            df_final.to_excel(stage1_filepath, index=False)
            logger.info(f"  Stage 1 Results saved to: {stage1_filepath}")
        
        success_count = len([r for r in results if r['First_Count'] is not None])
        failed_count  = len([r for r in results if r['First_Count'] is None])

        logger.info("="*80)
        logger.info(f"[OK] Stage 1 completed: {use_case}")
        logger.info(f"  Tables processed: {len(results)} | Success: {success_count} | Failed: {failed_count}")
        logger.info("="*80)

        return True   # ← was missing; None (falsy) was causing caller to log "Stage 1 failed"

    except Exception as e:
        logger.error(f"Stage 1 failed for {use_case}: {e}", exc_info=True)
        return False


# ============================================================================
# STAGE 2: CAPTURE SECOND COUNT
# ============================================================================

def capture_second_count_for_usecase(use_case):
    """
    STAGE 2: Capture second count for tables in specified use case
    
    Args:
        use_case: Use case name to process
    
    Returns:
        bool: True if successful, False otherwise
    """
    logger.info("="*80)
    logger.info(f"STAGE 2: CAPTURING SECOND COUNT - {use_case}")
    logger.info("="*80)
    
    try:
        stage1_filename = f"{use_case}_Stage1.xlsx"
        stage1_filepath = os.path.join(CONFIG_DIR, stage1_filename)
        
        stage2_filename = f"{use_case}_Stage2.xlsx"
        stage2_filepath = os.path.join(CONFIG_DIR, stage2_filename)

        # Check if first count file exists / load results
        if USE_S3:
            df_stage1 = read_excel_from_s3(stage1_filename)
            if df_stage1 is None:
                logger.error(f"Stage 1 file not found in S3 ({stage1_filename}). Run Stage 1 first.")
                return False
        else:
            if not os.path.exists(stage1_filepath):
                logger.error(f"Stage 1 file not found: {stage1_filepath}. Run Stage 1 first.")
                return False
            df_stage1 = pd.read_excel(stage1_filepath)
            
        # Note: We overwrite results for THIS use case, so no need to load existing
        # when paths are dynamic per use case.
        
        # Filter by use case
        df_usecase = df_stage1[df_stage1['replicated_jobs'] == use_case].copy()
        
        if len(df_usecase) == 0:
            logger.warning(f"No Stage 1 tables found for use case: {use_case}")
            return False
        
        logger.info(f"Processing {len(df_usecase)} tables for '{use_case}'")
        
        # Connect to database
        logger.info(f"[ACTION] Connecting to Snowflake to fetch second counts for {use_case}...")
        conn = db_connections("SNOWFLAKE", "Service")
        
        if not conn:
            logger.error("[ERROR] Database connection failed")
            return False
        
        logger.info("[OK] Database connection successful")
        
        # Process tables in batches
        success_count = 0
        failed_count = 0
        
        # Start a fresh Stage 2 results list
        stage2_results = []
        
        # Process all tables in one efficient lookup
        total_tables = len(df_usecase)
        logger.info(f"[ACTION] Processing all {total_tables} tables for second counts...")
        
        tables_to_query = []
        for _, row in df_usecase.iterrows():
            tables_to_query.append((row['Database_Name'], row['Schema_Name'], row['Table_Name']))
            
        all_counts = get_batch_table_counts(conn, tables_to_query)
        
        stage2_results = []
        success_count = 0
        failed_count = 0
        
        for _, row in df_usecase.iterrows():
            db, sc, tb = row['Database_Name'], row['Schema_Name'], row['Table_Name']
            count = all_counts.get((db, sc, tb)) if all_counts else None
            
            if count is None:
                count = get_table_count(conn, db, sc, tb)
                
            new_row = {
                'replicated_jobs': use_case,
                'Database_Name': db,
                'Schema_Name': sc,
                'Table_Name': tb,
                'Second_Count': count,
                'Second_Count_Timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }

            if count is not None:
                success_count += 1
            else:
                failed_count += 1
            
            stage2_results.append(new_row)
        
        # Close connection
        conn.close()
        
        # df_new_stage2 = pd.DataFrame(stage2_results)
        
        # Since paths are dynamic per use case, we just save the new results
        df_final_stage2 = pd.DataFrame(stage2_results)

        # Save stage 2 results to Excel (local or S3)
        if USE_S3:
            write_excel_to_s3(df_final_stage2, stage2_filename)
            logger.info(f"  Stage 2 Results saved to: s3://{S3_CONFIG['bucket']}/{_s3_key(stage2_filename)}")
        else:
            df_final_stage2.to_excel(stage2_filepath, index=False)
            logger.info(f"  Stage 2 Results saved to: {stage2_filepath}")
        
        logger.info("="*80)
        logger.info(f"[OK] Stage 2 completed: {use_case}")
        logger.info(f"  Tables processed: {total_tables} | Success: {success_count} | Failed: {failed_count}")
        logger.info("="*80)

        return True
        
    except Exception as e:
        logger.error(f"Stage 2 failed for {use_case}: {e}", exc_info=True)
        return False


# ============================================================================
# STAGE 3: VALIDATE COUNTS
# ============================================================================

def validate_counts_for_usecase(use_case):
    """
    STAGE 3: Validate counts for tables in specified use case
    
    Args:
        use_case: Use case name to process
    
    Returns:
        bool: True if successful, False otherwise
    """
    logger.info("="*80)
    logger.info(f"STAGE 3: VALIDATING COUNTS - {use_case}")
    logger.info("="*80)
    
    try:
        stage1_filename = f"{use_case}_Stage1.xlsx"
        stage1_filepath = os.path.join(CONFIG_DIR, stage1_filename)
        
        stage2_filename = f"{use_case}_Stage2.xlsx"
        stage2_filepath = os.path.join(CONFIG_DIR, stage2_filename)
        
        final_filename = f"Archive/{use_case}_Final.xlsx"
        final_filepath = os.path.join(CONFIG_DIR, final_filename)

        # Ensure Archive folder exists locally 
        if not os.path.exists(os.path.join(CONFIG_DIR, "Archive")):
            os.makedirs(os.path.join(CONFIG_DIR, "Archive"))


        # Load both stage 1 and stage 2 files 
        if USE_S3:
            df_stage1 = read_excel_from_s3(stage1_filename)
            df_stage2 = read_excel_from_s3(stage2_filename)
            # existing_final not needed as files are use-case specific
            
            if df_stage1 is None or df_stage2 is None:
                logger.error("Stage 1 or Stage 2 files not found in S3.")
                return False
        else:
            if not os.path.exists(stage1_filepath) or not os.path.exists(stage2_filepath):
                logger.error("Stage 1 or Stage 2 files not found locally. Run preceding stages.")
                return False
            df_stage1 = pd.read_excel(stage1_filepath)
            df_stage2 = pd.read_excel(stage2_filepath)

        # Prepare Stage 1: Only keep relevant columns and drop any existing Second_Count/Status from it
        cols_to_keep = ['replicated_jobs', 'Database_Name', 'Schema_Name', 'Table_Name', 'First_Count', 'First_Count_Timestamp']
        df_stage1_clean = df_stage1[[c for c in cols_to_keep if c in df_stage1.columns]].copy()
        
        # Merge them based on table keys
        df_merged = pd.merge(df_stage1_clean, df_stage2, on=['replicated_jobs', 'Database_Name', 'Schema_Name', 'Table_Name'], how='outer')
        
        # Filter by use case
        df_usecase = df_merged[df_merged['replicated_jobs'] == use_case].copy()
        
        if len(df_usecase) == 0:
            logger.warning(f"No tables found for use case: {use_case}")
            return False
            
        logger.info(f"Validating {len(df_usecase)} tables for '{use_case}'")
        
        passed_count = 0
        failed_count = 0
        error_count = 0
        
        # Validate each row
        logger.info(f"[ACTION] Comparing counts for {len(df_usecase)} tables...")
        validation_results = []
        for _, row in df_usecase.iterrows():
            table_name = row['Table_Name']
            first_count = row['First_Count'] if 'First_Count' in row else pd.NA
            second_count = row['Second_Count'] if 'Second_Count' in row else pd.NA
            
            logger.info(f"  Validating {table_name}: Yesterday={first_count} | Today={second_count}")
            
            status = 'PENDING'
            remarks = ''
            
            # Skip if either count is missing
            if pd.isna(first_count) or pd.isna(second_count):
                status = 'ERROR'
                remarks = 'Missing count data'
                error_count += 1
            elif int(second_count) == 0 and int(first_count) > 0:
                # FAILED: Table had data before but Second Count is zero now
                status = 'FAILED'
                remarks = f'Validation failed - Second Count is 0 but First Count was {int(first_count):,}'
                failed_count += 1
            else:
                # PASSED: Covers both counts equal, both zero, or second_count > 0
                status = 'PASSED'
                difference = int(second_count) - int(first_count)
                if int(first_count) == 0 and int(second_count) == 0:
                    remarks = 'Validation passed - Both counts are zero (normal)'
                elif difference >= 0:
                    remarks = f'Validation passed - Count increased/unchanged by {difference:,}'
                else:
                    remarks = f'Validation passed - Count changed by {difference:,}'
                passed_count += 1
                
            validation_results.append({
                'replicated_jobs': use_case,
                'Database_Name': row['Database_Name'],
                'Schema_Name': row['Schema_Name'],
                'Table_Name': row['Table_Name'],
                'First_Count': first_count,
                'Second_Count': second_count,
                'First_Count_Timestamp': row.get('First_Count_Timestamp', pd.NA),
                'Second_Count_Timestamp': row.get('Second_Count_Timestamp', pd.NA),
                'Status': status,
                'Remarks': remarks
            })
            
        df_new_validation = pd.DataFrame(validation_results)
        
        # Since paths are dynamic per use case, we just save the new results
        df_final_output = pd.DataFrame(validation_results)
            
        # Save updated results (local or S3)
        if USE_S3:
            write_excel_to_s3(df_final_output, final_filename)
        else:
            df_final_output.to_excel(final_filepath, index=False)
        
        logger.info("="*80)
        logger.info(f"[OK] Stage 3 completed: {use_case}")
        logger.info(f"  Tables validated: {len(df_usecase)} | Passed: {passed_count} | Failed: {failed_count} | Error: {error_count}")
        logger.info(f"  Results updated in: {final_filename}")
        logger.info("="*80)
        
        return True
        
    except Exception as e:
        logger.error(f"Stage 3 failed for {use_case}: {e}", exc_info=True)
        return False


# ============================================================================
# STAGE 4: CONSOLIDATE FAILED AND SEND EMAIL
# ============================================================================

def consolidate_failed_and_send_email_for_usecase(use_case, recipients):
    """
    STAGE 4: Consolidate failed validations and send an email notification.

    Logic:
      - If second_count == 0 AND first_count > 0  →  Status = 'FAILED'  →  email sent
      - If first_count == 0 AND second_count == 0 →  Status = 'PASSED'  →  email sent (normal run)
      - If second_count > 0 (regardless of first)  →  Status = 'PASSED'  →  email sent
      - Email is always sent on both PASS and FAIL.

    Recipients are passed IN from the schedule row ('email_recipients' column
    in Replicate_schedule_Config_test.xlsx), so each use case gets its OWN mailing list.

    Args:
        use_case:   Use case / replicated-job name to process
        recipients: list[str] — email addresses for THIS use case only

    Returns:
        bool: True if completed (with or without failures), False on error
    """
    logger.info("="*80)
    logger.info(f"STAGE 4: CONSOLIDATING FAILED RECORDS - {use_case}")
    logger.info("="*80)

    try:
        final_filename = f"Archive/{use_case}_Final.xlsx"
        final_filepath = os.path.join(CONFIG_DIR, final_filename)

        # ── Load final results (local or S3) ─────────────────────────────────────────
        if USE_S3:
            df_results = read_excel_from_s3(final_filename)
            if df_results is None:
                logger.error("Results file not found in S3. Run previous stages first.")
                return False
        else:
            if not os.path.exists(final_filepath):
                logger.error(f"Results file not found: {final_filepath}. Run previous stages first.")
                return False
            df_results = pd.read_excel(final_filepath)

        # ── Filter ─────────────────────────────────────────────────────────────
        df_usecase_total = df_results[df_results['replicated_jobs'] == use_case]
        df_failed = df_results[
            (df_results['replicated_jobs'] == use_case) &
            (df_results['Status'] == 'FAILED')
        ].copy()

        total  = len(df_usecase_total)
        failed = len(df_failed)
        passed = total - failed

        logger.info(f"Total: {total} | Passed: {passed} | Failed: {failed}")

        # ── Email notification triggers for BOTH Pass and Fail ──────────────────
        if failed == 0:
            logger.info(f"[OK] All {total} tables PASSED for '{use_case}'")
            status_to_send = 'PASSED'
            df_to_send = None
        else:
            logger.warning(f"VALIDATION FAILURES DETECTED for '{use_case}': {failed}/{total} tables FAILED")
            status_to_send = 'FAILED'
            df_to_send = df_failed

        # ── Send the email ────────────────────────────────────────────────────
        send_result = send_validation_notification(
            use_case=use_case,
            actual_count=passed,
            expected_count=total,
            failed_df=df_to_send,
            recipients=recipients,
            status=status_to_send
        )

        # CLEANUP: Delete the result files after email is sent
        # Now handles both local and S3 cleanup to avoid clutter
        stage1_filename = f"{use_case}_Stage1.xlsx"
        stage2_filename = f"{use_case}_Stage2.xlsx"
        stage1_filepath = os.path.join(CONFIG_DIR, stage1_filename)
        stage2_filepath = os.path.join(CONFIG_DIR, stage2_filename)
        
        # Local cleanup - Only delete Stage 1 and Stage 2 files
        for file_path in [stage1_filepath, stage2_filepath]:
            if os.path.exists(file_path):
                try:
                    os.remove(file_path)
                    logger.info(f"[CLEANUP] Deleted local temporary file: {file_path}")
                except Exception as e:
                    logger.warning(f"[CLEANUP] Could not delete local {file_path}: {e}")
        
        # S3 cleanup - Only delete Stage 1 and Stage 2 files
        if USE_S3:
            for s3_fname in [stage1_filename, stage2_filename]:
                try:
                    # Optional: only delete if they exist
                    delete_s3_file(s3_fname)
                    logger.info(f"[CLEANUP] Deleted S3 temporary file: {s3_fname}")
                except Exception as e:
                    logger.warning(f"[CLEANUP] Could not delete S3 {s3_fname}: {e}")

        return send_result

    except Exception as e:
        logger.error(f"Stage 4 failed for '{use_case}': {e}", exc_info=True)
        return False


# ============================================================================
# MAIN SCHEDULER FUNCTION
# ============================================================================

def run_scheduled_pipeline():
    """
    Main function to check schedules and run appropriate pipeline stages
    
    This function:
    1. Loads schedules from Excel file
    2. Gets current time
    3. Checks against all configured schedules
    4. Runs Stage 1 if current time matches first_run_time
    5. Runs Stages 2, 3, 4 if current time matches second_run_time
    """
    # Get current time in UTC to match the Excel schedule (e.g. 11:45 UTC)
    # Even if the machine is IST, this converts to UTC for the comparison.
    current_datetime = datetime.now(pytz.utc).replace(tzinfo=None)
    current_time_str = current_datetime.strftime("%H:%M")
    
    logger.info("="*80)
    logger.info("SCHEDULED PIPELINE RUNNER (UTC MODE)")
    logger.info(f"Current UTC Time: {current_datetime.strftime('%Y-%m-%d %H:%M:%S')}")
    logger.info("="*80)
    
    # Load schedules from Excel
    schedules = load_pipeline_schedules()
    
    if not schedules:
        logger.error("No schedules loaded. Check your configuration.")
        return
    
    logger.info(f"Checking {len(schedules)} configured schedule(s)...")
    
    executed_count = 0
    
    for schedule in schedules:
        use_case = schedule.get('replicated_jobs', '')
        first_run_time = str(schedule.get('job_1_time_utc', ''))
        second_run_time = str(schedule.get('job_2_time_utc', ''))
        
        logger.debug(f"Checking schedule (UTC): {use_case} | First: {first_run_time} | Second: {second_run_time}")
        
        # Check both times
        first_matches, first_diff = check_schedule_match(current_datetime, first_run_time)
        second_matches, second_diff = check_schedule_match(current_datetime, second_run_time)
        
        # Decide which stages to run. If both match, run Stage 1 then Stages 2-4.
        # This occurs if the two scheduled times are the same or very close.
        
        # 1. Check/Run Stage 1
        if first_matches:
            logger.info(f"[OK] MATCH! Running Stage 1 for: {use_case} (Diff: {first_diff} min)")

            # ── Open job-specific log identifying it by its start time ──
            job_handler, job_log_path = add_job_log_handler(use_case, "Stage1", first_run_time)

            success_s1 = capture_first_count_for_usecase(use_case)

            # ── Close & upload job-specific log ──
            remove_job_log_handler(job_handler, job_log_path)

            if success_s1:
                logger.info(f"[OK] Stage 1 completed successfully for: {use_case}")
                executed_count += 1
            else:
                logger.error(f"Stage 1 failed for: {use_case}")

        # 2. Check/Run Stages 2, 3, 4
        if second_matches:
            logger.info(f"[OK] MATCH! Running Stages 2, 3, 4 for: {use_case} (Diff: {second_diff} min)")

            # ── Open job-specific log, identifying it by its Stage 1 start time ──
            job_handler, job_log_path = add_job_log_handler(use_case, "Stage2-4", first_run_time)

            # Stage 2: Capture second count
            success_stage2 = capture_second_count_for_usecase(use_case)

            if not success_stage2:
                logger.error(f"Stage 2 failed for: {use_case}")
                remove_job_log_handler(job_handler, job_log_path)
                continue

            # Stage 3: Validate counts
            success_stage3 = validate_counts_for_usecase(use_case)

            if not success_stage3:
                logger.error(f"Stage 3 failed for: {use_case}")
                remove_job_log_handler(job_handler, job_log_path)
                continue

            # Stage 4: Consolidate and send email
            raw_emails = str(schedule.get('email_recipients', '')).strip()
            use_case_recipients = [
                e.strip() for e in raw_emails.split(',')
                if e.strip() and e.strip().lower() != 'nan'
            ]
            if not use_case_recipients:
                logger.debug(f"No specific recipients for '{use_case}', checking global list...")
                use_case_recipients = load_email_recipients()
                
            if not use_case_recipients:
                logger.warning(
                    f"No recipients found for '{use_case}' in any configuration. "
                    f"No email will be sent."
                )
            success_stage4 = consolidate_failed_and_send_email_for_usecase(use_case, use_case_recipients)

            # ── Close & upload job-specific log ──
            remove_job_log_handler(job_handler, job_log_path)

            if success_stage4:
                logger.info(f"[OK] Stages 2, 3, 4 completed successfully for: {use_case}")
                executed_count += 1
            else:
                logger.warning(f"Stages 2, 3, 4 completed with warnings for: {use_case}")
                executed_count += 1
    
    logger.info("="*80)
    if executed_count > 0:
        logger.info(f"[OK] PIPELINE EXECUTION COMPLETED - Executed {executed_count} schedule(s)")
        
        # Check if all scheduled jobs are completed and cleanup if needed
        cleanup_performed = check_and_cleanup_completed_jobs()
        
        if not cleanup_performed:
            logger.debug("Cleanup not performed - waiting for remaining jobs to complete")
    else:
        logger.info(f"NO SCHEDULES MATCHED CURRENT TIME: {current_time_str}")
        logger.info("Next scheduled times:")
        for schedule in schedules:
            logger.info(f"  - {schedule.get('replicated_jobs', 'Unknown')}: {schedule.get('job_1_time_ist', 'N/A')} (Stage 1), {schedule.get('job_2_time_ist', 'N/A')} (Stages 2-4)")
    logger.info("="*80)

    # ── Upload log file to S3 (AWS mode only) ──
    # Only upload if we actually did something, to avoid cluttering S3 with empty minute-by-minute logs
    if executed_count > 0:
        upload_log_to_s3()

    return executed_count


# ============================================================================
# CONTINUOUS MODE - RUNS FOREVER
# ============================================================================

def run_continuous_mode(check_interval_seconds=60):
    """
    Run the pipeline continuously in a loop
    
    This function runs forever, checking schedules every minute and executing
    when the current time matches a scheduled time.
    
    Args:
        check_interval_seconds: How often to check schedules (default: 60 seconds)
    """
    logger.info("="*80)
    logger.info("CONTINUOUS MODE - SCHEDULED PIPELINE RUNNER")
    logger.info(f"Check Interval: Every {check_interval_seconds} seconds")
    logger.info("Press Ctrl+C to stop")
    logger.info("="*80)
    
    # Track executed schedules to avoid running the same schedule multiple times
    # Format: {(use_case, stage_type, date, hour, minute): True}
    executed_today = {}
    last_date = datetime.now().date()
    
    try:
        while True:
            # Force current time to UTC to match Excel (regardless of local machine time)
            current_datetime = datetime.now(pytz.utc).replace(tzinfo=None)
            current_date = current_datetime.date()
            
            # Reset executed tracking at midnight
            if current_date != last_date:
                logger.info("="*80)
                logger.info(f"NEW DAY: {current_date}")
                logger.info("Resetting execution tracking...")
                logger.info("="*80)
                executed_today = {}
                last_date = current_date
            
            # Load schedules from Excel (allows hot-reload if file is updated)
            schedules = load_pipeline_schedules()
            
            if not schedules:
                logger.warning(f"[{current_datetime.strftime('%Y-%m-%d %H:%M:%S')}] No schedules loaded. Waiting {check_interval_seconds}s...")
                time_module.sleep(check_interval_seconds)
                continue
            
            # Check each schedule
            for schedule in schedules:
                use_case = schedule.get('replicated_jobs', '')
                first_run_time = str(schedule.get('job_1_time_utc', ''))
                second_run_time = str(schedule.get('job_2_time_utc', ''))
                
                # Check for Stage 1 execution (check independently, not elif)
                matched_stage1, diff1 = check_schedule_match(current_datetime, first_run_time)
                if matched_stage1:
                    # Key based on the SCHEDULED time, not current time, to prevent double-runs
                    execution_key = (use_case, 'stage1', current_date, first_run_time)
                    
                    if execution_key not in executed_today:
                        logger.info(f"[{current_datetime.strftime('%Y-%m-%d %H:%M:%S')}] [OK] MATCH! Running Stage 1 for: {use_case}")
                        
                        success = capture_first_count_for_usecase(use_case)
                        
                        if success:
                            logger.info(f"[OK] Stage 1 completed for: {use_case}")
                            executed_today[execution_key] = True
                        else:
                            logger.error(f"Stage 1 failed for: {use_case}")
                
                # Check for Stages 2, 3, 4 execution (separate if, not elif - allows both to run)
                matched_stage2, diff2 = check_schedule_match(current_datetime, second_run_time)
                if matched_stage2:
                    # Key based on the SCHEDULED time, not current time, to prevent double-runs
                    execution_key = (use_case, 'stages234', current_date, second_run_time)
                    
                    if execution_key not in executed_today:
                        logger.info(f"[{current_datetime.strftime('%Y-%m-%d %H:%M:%S')}] [OK] MATCH! Running Stages 2, 3, 4 for: {use_case}")
                        
                        # Stage 2
                        success_stage2 = capture_second_count_for_usecase(use_case)
                        if not success_stage2:
                            logger.error(f"Stage 2 failed for: {use_case}")
                            continue
                        
                        # Stage 3
                        success_stage3 = validate_counts_for_usecase(use_case)
                        if not success_stage3:
                            logger.error(f"Stage 3 failed for: {use_case}")
                            continue
                        
                        # Stage 4 - Parse recipients
                        raw_emails = str(schedule.get('email_recipients', '')).strip()
                        use_case_recipients = [
                            e.strip() for e in raw_emails.split(',')
                            if e.strip() and e.strip().lower() != 'nan'
                        ]
                        if not use_case_recipients:
                            use_case_recipients = load_email_recipients()

                        success_stage4 = consolidate_failed_and_send_email_for_usecase(use_case, use_case_recipients)
                        
                        if success_stage4:
                            logger.info(f"[OK] Stages 2, 3, 4 completed for: {use_case}")
                            executed_today[execution_key] = True
                        else:
                            logger.warning(f"Stages 2, 3, 4 completed with warnings for: {use_case}")
                            executed_today[execution_key] = True
            
            # After processing all schedules, check if all jobs are completed and cleanup if needed
            check_and_cleanup_completed_jobs()
            
            # Wait before next check
            logger.debug(f"[{current_datetime.strftime('%Y-%m-%d %H:%M:%S')}] Waiting {check_interval_seconds}s for next check...")
            time_module.sleep(check_interval_seconds)
    
    except KeyboardInterrupt:
        logger.info("="*80)
        logger.info("CONTINUOUS MODE STOPPED BY USER")
        logger.info("="*80)
        sys.exit(0)



# ============================================================================
# MAIN ENTRY POINT
# ============================================================================

if __name__ == "__main__":
    """
    AWS Glue / Cron entry point.

    The script runs once and exits.
    EventBridge triggers it every minute — it checks the schedule Excel on S3,
    runs the matching stage (1 or 2-3-4) for any use case whose time matches,
    then exits. Logs are uploaded to S3 automatically after each run.

    EventBridge cron: cron(* * * * ? *)   ← every minute (UTC)
    """
    run_scheduled_pipeline()
