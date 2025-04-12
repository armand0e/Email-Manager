# AI Email Manager

This application connects to your Gmail or Outlook account, fetches your emails,
and uses NLP (NLTK, spaCy) to analyze, categorize, and prioritize them.

## Features

- Connects to Gmail and Outlook using IMAP/Exchange.
- Fetches emails with pagination ("Load More").
- Analyzes email content (subject, body) for category detection (Work, Personal, Newsletter, Notification, Other).
- Scores email priority based on sender, keywords, time sensitivity (implementation in `priority_scorer.py`).
- Displays emails in a Streamlit interface, sorted by priority (implicitly, more work needed here).
- Allows overriding calculated priority (High, Medium, Low).
- Provides buttons to open emails in the native web client (Gmail/Outlook popout) or start a reply.
- Basic filtering by priority and category.
- Session persistence for priority overrides (saved to `.session_data.json`).

## Setup

1.  **Clone the repository:**
    ```bash
    git clone <repository-url>
    cd Email-Manager
    ```

2.  **Create a virtual environment (recommended):**
    ```bash
    python -m venv venv
    # Activate the environment
    # Windows
    .\venv\Scripts\activate
    # macOS/Linux
    source venv/bin/activate
    ```

3.  **Install dependencies:**
    ```bash
    pip install -r requirements.txt
    ```

4.  **Download NLP models:**
    *   **spaCy English model:**
        ```bash
        python -m spacy download en_core_web_sm
        ```
    *   **NLTK data:** The application attempts to download `punkt`, `stopwords`, and `wordnet` automatically on first run if they are missing. Ensure you have an internet connection.

5.  **Environment Variables (Optional):**
    Create a `.env` file in the root directory (copy from `.env.example`). Currently, no specific API keys are required here as credentials are entered in the UI, but this file can be used for future configuration.

## Running the Application

```bash
streamlit run app.py
```

Open your browser to the local URL provided by Streamlit (usually http://localhost:8501).

## Usage

1.  Select either Gmail or Outlook.
2.  Enter your email address.
3.  Enter your password. **Important:** If you use 2-Factor Authentication (2FA), you **must** generate and use an **App Password** specific to this application. Do not use your regular account password.
4.  Click "Connect".
5.  Emails will be fetched, analyzed, and displayed.
6.  Use the sidebar filters or the "Load More" button as needed.
7.  Expand emails to view snippets and use action buttons (Reply, Open, Change Priority).

## Future Improvements / TODO

- [ ] Implement actual Archive and Mark as Read functionality (requires write permissions).
- [ ] Improve priority scoring logic.
- [ ] Enhance category detection (more keywords, ML model?).
- [ ] Add explicit sorting options.
- [ ] Improve UI responsiveness and state management during actions.
- [ ] Implement OAuth2 for authentication instead of password entry.
- [ ] Add unit and integration tests.
- [ ] More robust handling of different email formats/encodings.
- [ ] Better configuration management (e.g., for keywords, scoring weights).
- [ ] Asynchronous email fetching/processing.

## Security

- Your email credentials are only used for the current session
- No credentials are stored permanently
- You can disconnect at any time
- No email content is stored locally

## Requirements

- Python 3.8+
- Gmail or Outlook account
- Internet connection
