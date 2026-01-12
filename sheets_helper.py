"""
Google Sheets helper functions for reading and writing data
"""
import gspread
from google.oauth2 import service_account
import os
from typing import List, Tuple, Optional
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Path to Google Sheets API credentials JSON file
# Load from .env file, fallback to default filename if not set
SHEETS_CREDENTIALS_PATH = os.getenv('GOOGLE_SHEETS_CREDENTIALS_PATH', 'groovy-electron-478008-k6-a6eb0ee3e332.json')

def _validate_and_fix_credentials_file(file_path: str) -> dict:
    """
    Validate and fix credentials file if needed
    Returns the credentials dict
    """
    import json
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            creds_data = json.load(f)
        
        # Fix private key if it has literal \n instead of actual newlines
        if 'private_key' in creds_data:
            private_key = creds_data['private_key']
            # Check if it has literal \n (escaped) but not actual newlines
            if '\\n' in private_key and private_key.count('\n') < 5:
                # Replace literal \n with actual newlines
                creds_data['private_key'] = private_key.replace('\\n', '\n')
                # Optionally save the fixed version back (commented out for safety)
                # with open(file_path, 'w', encoding='utf-8') as f:
                #     json.dump(creds_data, f, indent=2)
        
        return creds_data
    except json.JSONDecodeError as e:
        raise Exception(f"認証情報ファイルのJSON形式が無効です: {str(e)}")
    except Exception as e:
        raise Exception(f"認証情報ファイルの読み込みエラー: {str(e)}")

def get_sheets_client():
    """
    Get authenticated Google Sheets client
    Returns: gspread.Client instance
    """
    if not os.path.exists(SHEETS_CREDENTIALS_PATH):
        raise FileNotFoundError(f"Google Sheets credentials file not found: {SHEETS_CREDENTIALS_PATH}")
    
    scope = [
        'https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive'
    ]
    
    try:
        # First, validate and fix credentials file if needed
        creds_info = _validate_and_fix_credentials_file(SHEETS_CREDENTIALS_PATH)
        
        # Try loading from dict first (more reliable for handling encoding issues)
        try:
            creds = service_account.Credentials.from_service_account_info(
                creds_info,
                scopes=scope
            )
        except Exception:
            # Fallback to file loading
            creds = service_account.Credentials.from_service_account_file(
                SHEETS_CREDENTIALS_PATH,
                scopes=scope
            )
    except Exception as e:
        error_msg = str(e)
        if 'invalid_grant' in error_msg.lower() or 'jwt' in error_msg.lower():
            raise Exception(
                f"認証エラー: JWT署名が無効です。\n\n"
                f"このエラーは通常、認証情報ファイルが破損しているか、無効な秘密鍵が含まれている場合に発生します。\n\n"
                f"解決方法:\n"
                f"1. Google Cloud Console (https://console.cloud.google.com/) にアクセス\n"
                f"2. プロジェクト 'groovy-electron-478008-k6' を選択\n"
                f"3. 'IAM & Admin' → 'Service Accounts' に移動\n"
                f"4. サービスアカウント 'pokemon-sheet@groovy-electron-478008-k6.iam.gserviceaccount.com' を選択\n"
                f"5. 'Keys' タブ → 'Add Key' → 'Create new key' → 'JSON' を選択\n"
                f"6. ダウンロードした新しいJSONファイルで '{SHEETS_CREDENTIALS_PATH}' を置き換えてください\n\n"
                f"元のエラー: {error_msg}"
            )
        else:
            raise Exception(
                f"認証情報の読み込みに失敗しました: {error_msg}\n\n"
                f"認証情報ファイル '{SHEETS_CREDENTIALS_PATH}' が有効であることを確認してください。\n"
                f"必要に応じて、Google Cloud Consoleから新しい認証情報をダウンロードしてください。"
            ) from e
    
    try:
        client = gspread.authorize(creds)
        return client
    except Exception as e:
        error_msg = str(e)
        if 'invalid_grant' in error_msg.lower() or 'jwt' in error_msg.lower():
            raise Exception(
                f"認証エラー: JWT署名が無効です。\n\n"
                f"このエラーは通常、認証情報ファイルが破損しているか、無効な秘密鍵が含まれている場合に発生します。\n\n"
                f"解決方法:\n"
                f"1. Google Cloud Console (https://console.cloud.google.com/) にアクセス\n"
                f"2. プロジェクト 'groovy-electron-478008-k6' を選択\n"
                f"3. 'IAM & Admin' → 'Service Accounts' に移動\n"
                f"4. サービスアカウント 'pokemon-sheet@groovy-electron-478008-k6.iam.gserviceaccount.com' を選択\n"
                f"5. 'Keys' タブ → 'Add Key' → 'Create new key' → 'JSON' を選択\n"
                f"6. ダウンロードした新しいJSONファイルで '{SHEETS_CREDENTIALS_PATH}' を置き換えてください\n\n"
                f"元のエラー: {error_msg}"
            )
        else:
            raise

def extract_spreadsheet_id(spreadsheet_input: str) -> str:
    """
    Extract spreadsheet ID from URL or use as-is if it's already an ID
    
    Examples:
        - "https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit"
        - "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms"
    
    Returns: Spreadsheet ID string
    """
    spreadsheet_input = spreadsheet_input.strip()
    
    # Check if it's a URL
    if 'docs.google.com/spreadsheets/d/' in spreadsheet_input:
        # Extract ID from URL
        parts = spreadsheet_input.split('/d/')
        if len(parts) > 1:
            spreadsheet_id = parts[1].split('/')[0]
            return spreadsheet_id
    
    # If it's not a URL, assume it's already an ID
    return spreadsheet_input

def read_sheets_data(spreadsheet_id: str, worksheet_name: str = None, start_row: int = None, end_row: int = None) -> Tuple[List[Tuple], int, int]:
    """
    Read data from Google Sheets with optional row range
    
    Column mapping:
        - Column A (index 0): Email address (ログイン情報)
        - Column B (index 1): Password (パスワード)
        - Column C (index 2): Status (状態) - "成功"の場合はスキップ
        - Column D (index 3): Detailed message (具体的な状態) - 読み取り専用
        - Column E (index 4): Timestamp (最終進行時間) - 読み取り専用
    
    Args:
        spreadsheet_id: Google Spreadsheet ID or URL
        worksheet_name: Name of the worksheet (default: first sheet)
        start_row: Start row number (1-based, inclusive). If None, starts from row 1.
        end_row: End row number (1-based, inclusive). If None, processes all rows.
    
    Returns:
        Tuple of (data_rows, total_email_count, skipped_count) where:
        - data_rows: List of (row_number, email, password) tuples for rows to process
        - total_email_count: Total number of rows with email addresses in the specified range
        - skipped_count: Number of rows skipped (already "成功" in column C)
    """
    try:
        client = get_sheets_client()
        spreadsheet_id = extract_spreadsheet_id(spreadsheet_id)
        spreadsheet = client.open_by_key(spreadsheet_id)
        
        # Get worksheet (same worksheet used for writing)
        if worksheet_name:
            worksheet = spreadsheet.worksheet(worksheet_name)
        else:
            worksheet = spreadsheet.sheet1  # Use first sheet
        
        # Get all values from the same spreadsheet
        all_values = worksheet.get_all_values()
        
        data_rows = []
        skipped_count = 0
        total_email_count = 0
        
        # Determine row range
        if start_row is not None and end_row is not None:
            if start_row > end_row:
                raise ValueError(f"Start row ({start_row}) must be less than or equal to end row ({end_row})")
            if start_row < 1:
                raise ValueError(f"Start row must be at least 1, got {start_row}")
        
        # Process each row (skip header row if exists)
        for i, row in enumerate(all_values, start=1):
            # Apply row range filter if specified
            if start_row is not None and i < start_row:
                continue  # Skip rows before start_row
            if end_row is not None and i > end_row:
                break  # Stop processing after end_row
            
            if not row or not row[0] or not row[0].strip():
                continue  # Skip empty rows
            
            # Column A: Email address (ログイン情報)
            email = row[0].strip()
            
            # Column B: Password (パスワード)
            password = row[1].strip() if len(row) > 1 and row[1] else None
            
            if not email:
                continue
            
            total_email_count += 1
            
            # Column C: Check for "成功" status (状態)
            column_c_value = None
            if len(row) > 2 and row[2]:
                column_c_value = str(row[2]).strip()
            
            # Skip rows that are already "成功"
            if column_c_value == "成功":
                skipped_count += 1
            else:
                # Add to processing list: (row_number, email, password)
                data_rows.append((i, email, password))
        
        return data_rows, total_email_count, skipped_count
    
    except Exception as e:
        raise Exception(f"Error reading Google Sheets: {str(e)}")

def write_sheets_result(spreadsheet_id: str, row_number: int, status: str, message: str, timestamp: str, worksheet_name: str = None):
    """
    Write result to Google Sheets (same spreadsheet used for reading)
    
    Args:
        spreadsheet_id: Google Spreadsheet ID or URL (same as used for reading)
        row_number: Row number (1-based)
        status: Status to write to column C (状態)
        message: Message to write to column D (具体的な状態)
        timestamp: Timestamp to write to column E (最終進行時間)
        worksheet_name: Name of the worksheet (default: first sheet, same as used for reading)
    
    Column mapping:
        - Column A (index 0): Email address (読み込み専用)
        - Column B (index 1): Password (読み込み専用)
        - Column C (index 2): Status (書き込み: 成功/失敗)
        - Column D (index 3): Detailed message (書き込み: 具体的な状態)
        - Column E (index 4): Timestamp (書き込み: 最終進行時間)
    """
    try:
        client = get_sheets_client()
        spreadsheet_id = extract_spreadsheet_id(spreadsheet_id)
        spreadsheet = client.open_by_key(spreadsheet_id)
        
        # Get worksheet (same worksheet used for reading)
        if worksheet_name:
            worksheet = spreadsheet.worksheet(worksheet_name)
        else:
            worksheet = spreadsheet.sheet1  # Use first sheet
        
        # Use batch update for better performance and atomicity
        # Update columns C, D, E in a single batch operation
        range_name = f"C{row_number}:E{row_number}"
        values = [[status, message, timestamp]]
        worksheet.update(range_name, values, value_input_option='USER_ENTERED')
        
    except Exception as e:
        raise Exception(f"Error writing to Google Sheets (row {row_number}): {str(e)}")

def _get_service_account_email() -> str:
    """Get service account email from credentials file"""
    try:
        import json
        if os.path.exists(SHEETS_CREDENTIALS_PATH):
            with open(SHEETS_CREDENTIALS_PATH, 'r', encoding='utf-8') as f:
                creds_data = json.load(f)
                return creds_data.get('client_email', 'サービスアカウント')
    except:
        pass
    return 'サービスアカウント'

def check_sheets_access(spreadsheet_id: str, worksheet_name: str = None) -> Tuple[bool, str]:
    """
    Check if we can access the spreadsheet
    
    Args:
        spreadsheet_id: Google Spreadsheet ID or URL
        worksheet_name: Name of the worksheet (default: first sheet)
    
    Returns:
        Tuple of (is_accessible: bool, error_message: str)
        If accessible, returns (True, "")
        If not accessible, returns (False, error_message)
    """
    try:
        client = get_sheets_client()
        spreadsheet_id = extract_spreadsheet_id(spreadsheet_id)
        spreadsheet = client.open_by_key(spreadsheet_id)
        
        if worksheet_name:
            worksheet = spreadsheet.worksheet(worksheet_name)
        else:
            worksheet = spreadsheet.sheet1
        
        # Try to read first row to verify access
        worksheet.get('A1')
        return True, ""
    except FileNotFoundError as e:
        return False, f"認証情報ファイルが見つかりません: {str(e)}"
    except gspread.exceptions.SpreadsheetNotFound:
        return False, f"スプレッドシートが見つかりません。ID '{spreadsheet_id}' が正しいか確認してください。"
    except gspread.exceptions.APIError as e:
        error_code = e.response.status_code if hasattr(e, 'response') else 'unknown'
        if error_code == 403:
            service_account_email = _get_service_account_email()
            return False, f"アクセスが拒否されました (403)。スプレッドシートにサービスアカウント '{service_account_email}' を共有してください。"
        elif error_code == 404:
            return False, f"スプレッドシートが見つかりません (404)。ID '{spreadsheet_id}' が正しいか確認してください。"
        else:
            return False, f"Google Sheets API エラー ({error_code}): {str(e)}"
    except Exception as e:
        error_msg = str(e)
        # Check for JWT signature errors
        if "invalid_grant" in error_msg.lower() or "jwt" in error_msg.lower() or "invalid jwt signature" in error_msg.lower():
            return False, (
                f"認証エラー: JWT署名が無効です。\n\n"
                f"認証情報ファイルが破損している可能性があります。\n"
                f"Google Cloud Consoleから新しい認証情報をダウンロードして置き換えてください。\n\n"
                f"詳細: {error_msg}"
            )
        # Check for common error patterns
        elif "PERMISSION_DENIED" in error_msg or "permission" in error_msg.lower():
            service_account_email = _get_service_account_email()
            return False, f"アクセス権限がありません。スプレッドシートにサービスアカウント '{service_account_email}' を共有してください。"
        elif "NOT_FOUND" in error_msg or "not found" in error_msg.lower():
            return False, f"スプレッドシートが見つかりません。ID '{spreadsheet_id}' が正しいか確認してください。"
        else:
            return False, f"エラー: {error_msg}"
