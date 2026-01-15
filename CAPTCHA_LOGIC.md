# CAPTCHA Logic Documentation

This document explains the CAPTCHA solving logic used in the Pokemon Auto Lottery bot.

## Overview

The bot uses the **2Captcha API** service to solve **reCAPTCHA v3 Enterprise** challenges. The logic is implemented using the modern JSON-based API with the following main functions:

1. `solve_recaptcha()` - Core CAPTCHA solving function (uses new 2Captcha API)
2. `extract_recaptcha_site_key()` - Extracts reCAPTCHA site key from page
3. `extract_recaptcha_action()` - Extracts pageAction parameter if available
4. `_attempt_login_with_captcha()` - Handles CAPTCHA during login
5. `_check_and_solve_captcha_on_apply_page()` - Handles CAPTCHA on the apply page

## Configuration

The CAPTCHA API key is loaded from environment variables:
```python
CAPTCHA_API_KEY = os.getenv("CAPTCHA_API_KEY")
```

## Pokemon Center reCAPTCHA Details

- **Type:** reCAPTCHA v3 Enterprise
- **Site Key:** `6Le9HlYqAAAAAJQtQcq3V_tdd73twiM4Rm2wUvn9`
- **Domain:** `https://www.pokemoncenter-online.com/`
- **API Domain:** `www.google.com` (uses `recaptcha/enterprise.js`)

## Core Function: `solve_recaptcha()`

**Location:** Lines 297+ in `bot.py`

**Purpose:** Solves reCAPTCHA v3 Enterprise using 2Captcha API (new JSON-based API)

**Parameters:**
- `site_key` (str): The reCAPTCHA site key found on the page
- `url` (str): The full URL of target web page
- `driver` (WebDriver, optional): Selenium driver to extract pageAction
- `max_retries` (int): Maximum number of retry attempts (default: 5)
- `min_score` (float): Required score value: 0.3, 0.7, or 0.9 (default: 0.3)
- `page_action` (str, optional): Action parameter value

**Process Flow:**

1. **Extract pageAction (if not provided):**
   - Uses `extract_recaptcha_action()` to find action from page
   - Looks for `data-action` attribute or `grecaptcha.execute` calls

2. **Submit Task to 2Captcha:**
   - Uses POST request to `https://api.2captcha.com/createTask`
   - Sends JSON payload with task type `RecaptchaV3TaskProxyless`
   - Task includes:
     - `websiteURL`: Current page URL
     - `websiteKey`: reCAPTCHA site key
     - `minScore`: Required score (0.3, 0.7, or 0.9)
     - `isEnterprise`: `true` (Pokemon Center uses Enterprise)
     - `apiDomain`: `www.google.com`
     - `pageAction`: Extracted action (if available)

3. **Get Task ID:**
   - Extracts `taskId` from JSON response
   - Handles errors with retry logic

4. **Poll for Solution:**
   - Polls `https://api.2captcha.com/getTaskResult` every 5 seconds
   - Maximum 60 polls (5 minutes total)
   - Handles response states:
     - `processing`: Continue waiting
     - `ready`: Extract solution token from `solution.gRecaptchaResponse` or `solution.token`
     - `ERROR_CAPTCHA_UNSOLVABLE`: Retry with new attempt
     - Other errors: Retry or raise exception

5. **Error Handling:**
   - Retries on errors up to `max_retries` times
   - Raises exception if all retries fail
   - Respects stop signals from user

**Returns:** CAPTCHA solution token (string)

**Example Usage:**
```python
site_key = "6Le9HlYqAAAAAJQtQcq3V_tdd73twiM4Rm2wUvn9"
url = driver.current_url
solution = solve_recaptcha(site_key, url, driver=driver, min_score=0.3)
```

## Helper Functions

### `extract_recaptcha_site_key(driver)`

Extracts reCAPTCHA site key using multiple methods:
1. Checks for Pokemon Center specific site key
2. Extracts from `enterprise.js` or `api.js` script URL (`render=` parameter)
3. Looks for `data-sitekey` attribute
4. Falls back to regex pattern `6Le[a-zA-Z0-9_-]+`

### `extract_recaptcha_action(driver)`

Extracts `pageAction` parameter:
1. Checks `data-action` attribute in reCAPTCHA div
2. Searches for `grecaptcha.execute` or `grecaptcha.enterprise.execute` calls in scripts

## Login CAPTCHA: `_attempt_login_with_captcha()`

**Location:** Lines 512-645 in `bot.py`

**Purpose:** Handles CAPTCHA during the login process

**Process:**

1. **Load Login Page:**
   - Navigates to the login URL
   - Waits for page to load

2. **Detect CAPTCHA:**
   - Searches page source for reCAPTCHA site key using regex: `r'6Le[a-zA-Z0-9_-]+'`
   - Site keys typically start with "6Le"

3. **Solve CAPTCHA (if found):**
   - Calls `solve_recaptcha()` to get solution
   - Injects solution into the page using JavaScript:
     - Sets value of `g-recaptcha-response` textarea elements
     - Triggers reCAPTCHA callbacks if available
     - Handles `___grecaptcha_cfg.clients` for multiple CAPTCHA instances

4. **Continue Login:**
   - Enters email and password
   - Clicks login button
   - Checks login status

**Returns:** Tuple `(success, needs_retry)`
- `success`: True if login successful (redirected away from login page)
- `needs_retry`: True if authentication failed and should retry

## Apply Page CAPTCHA: `_check_and_solve_captcha_on_apply_page()`

**Location:** Lines 1297-1359 in `bot.py`

**Purpose:** Checks and solves CAPTCHA on the lottery apply page

**Process:**

1. **Verify Page:**
   - Checks if currently on `apply.html` page
   - Returns `False` if not on apply page

2. **Detect CAPTCHA:**
   - Searches page source for reCAPTCHA site key
   - Same regex pattern as login: `r'6Le[a-zA-Z0-9_-]+'`

3. **Solve CAPTCHA (if found):**
   - Calls `solve_recaptcha()` to get solution
   - Injects solution using the same JavaScript injection method
   - Waits 2 seconds for CAPTCHA to be processed

**Returns:** Boolean
- `True`: CAPTCHA was found and solved
- `False`: No CAPTCHA found or error occurred

**Usage Locations:**
- Called before starting lottery processing (line 1413)
- Called before processing each lottery entry (line 2206)
- Called before submitting application (line 2362)

## JavaScript Injection Method

The CAPTCHA solution is injected into the page using enhanced JavaScript for reCAPTCHA v3 Enterprise:

```javascript
var token = "{captcha_solution}";

// Set value in g-recaptcha-response textarea
var textareas = document.getElementsByName("g-recaptcha-response");
for (var i = 0; i < textareas.length; i++) {
    textareas[i].value = token;
}

// Also set in hidden input fields
var inputs = document.querySelectorAll('input[name="g-recaptcha-response"]');
for (var i = 0; i < inputs.length; i++) {
    inputs[i].value = token;
}

// Trigger callbacks for reCAPTCHA v3 Enterprise
if (typeof ___grecaptcha_cfg !== 'undefined') {
    Object.keys(___grecaptcha_cfg.clients).forEach(function(key) {
        var client = ___grecaptcha_cfg.clients[key];
        if (client && client.callback) {
            try {
                client.callback(token);
            } catch(e) {
                console.log("Callback error:", e);
            }
        }
    });
}

// Trigger grecaptcha.enterprise callback if available
if (typeof grecaptcha !== 'undefined' && grecaptcha.enterprise) {
    try {
        if (typeof window.grecaptchaCallbacks !== 'undefined') {
            for (var i = 0; i < window.grecaptchaCallbacks.length; i++) {
                try {
                    window.grecaptchaCallbacks[i](token);
                } catch(e) {
                    console.log("Callback error:", e);
                }
            }
        }
    } catch(e) {
        console.log("Enterprise callback error:", e);
    }
}

// Dispatch input events to notify the page
textareas = document.getElementsByName("g-recaptcha-response");
for (var i = 0; i < textareas.length; i++) {
    var event = new Event('input', { bubbles: true });
    textareas[i].dispatchEvent(event);
    var changeEvent = new Event('change', { bubbles: true });
    textareas[i].dispatchEvent(changeEvent);
}
```

This enhanced method:
1. Sets value in all `g-recaptcha-response` textarea and input elements
2. Triggers reCAPTCHA v3 Enterprise callbacks via `___grecaptcha_cfg.clients`
3. Handles `grecaptcha.enterprise` callbacks if available
4. Dispatches input/change events to notify the page of token updates
5. Includes error handling for callback failures

## Key Features

- **Automatic Detection:** Searches for CAPTCHA site keys in page source
- **Retry Logic:** Handles failures with configurable retries
- **Stop Signal Support:** Respects user stop requests during solving
- **Multiple Locations:** Handles CAPTCHA on login and apply pages
- **Invisible reCAPTCHA:** Configured for invisible reCAPTCHA challenges
- **Error Handling:** Comprehensive error handling with logging

## Dependencies

- `requests` library for API calls
- `re` module for regex pattern matching
- Selenium WebDriver for page interaction
- Environment variable `CAPTCHA_API_KEY` must be set

## Notes

- **API Version:** Uses the new 2Captcha JSON-based API (not the legacy GET-based API)
- **reCAPTCHA Type:** Configured for reCAPTCHA v3 Enterprise
- **Enterprise Support:** `isEnterprise: true` is set for Pokemon Center
- **Minimum Score:** Default is 0.3 (can be 0.3, 0.7, or 0.9)
- **Maximum Wait Time:** 300 seconds (60 polls × 5 seconds)
- **PageAction:** Automatically extracted from page if available
- **Site Key Detection:** Multiple methods including Pokemon Center specific key
- **Stop Support:** CAPTCHA solving can be stopped by user at any time
- **Logging:** All CAPTCHA operations are logged with emoji indicators for easy tracking

## API Endpoints

- **Create Task:** `https://api.2captcha.com/createTask` (POST, JSON)
- **Get Result:** `https://api.2captcha.com/getTaskResult` (POST, JSON)

## Task Type

- **Type:** `RecaptchaV3TaskProxyless`
- **Required Fields:**
  - `websiteURL`: Full URL of target page
  - `websiteKey`: reCAPTCHA site key
  - `minScore`: 0.3, 0.7, or 0.9
- **Optional Fields:**
  - `pageAction`: Action parameter value
  - `isEnterprise`: Boolean (true for Pokemon Center)
  - `apiDomain`: "www.google.com" or "www.recaptcha.net"
