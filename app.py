import os
import streamlit as st
from dotenv import load_dotenv
from email_analyzer import EmailAnalyzer
from email_client import EmailClient
from priority_scorer import PriorityScorer
import webbrowser
import urllib.parse
import time
import json
from datetime import datetime, timedelta
import re # For email validation
import traceback # For detailed error logging
import logging # Standard Python logging

# --- Setup Logging --- #
log_file = 'app.log'
# Basic configuration: Log to file and console
logging.basicConfig(
    level=logging.INFO, # Set default level (DEBUG, INFO, WARNING, ERROR, CRITICAL)
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file), # Log to a file
        logging.StreamHandler() # Log to console (Streamlit catches this)
    ]
)
logger = logging.getLogger(__name__) # Get logger for this module

# Load environment variables
load_dotenv()
logger.info("Application started, environment variables loaded.")

# Configure Streamlit page
st.set_page_config(
    page_title="AI Email Manager", # Updated title
    page_icon="📧",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
    <style>
    .stButton button {
        width: 100%;
        border-radius: 5px;
        height: 2.5em;
    }
    .email-header {
        font-size: 1.2em;
        color: #0066cc;
    }
    .priority-high {
        color: #ff4b4b;
        font-weight: bold;
    }
    .priority-medium {
        color: #ffa500;
        font-weight: bold;
    }
    .priority-low {
        color: #00cc00;
        font-weight: bold;
    }
    .email-metadata {
        color: #666;
        font-size: 0.9em;
    }
    .email-snippet {
        margin-top: 10px;
        margin-bottom: 10px;
    }
    /* Popovers can be styled if needed, but avoid max_width here */
    </style>
""", unsafe_allow_html=True)

# --- Session State Management --- #

def save_session_state():
    """Save important session state data (like priority overrides) to a local file."""
    if 'priority_overrides' not in st.session_state:
        logger.warning("Attempted to save session state before priority_overrides was initialized.")
        return # Don't save if state isn't ready
        
    try:
        state_data = {
            'priority_overrides': st.session_state.get('priority_overrides', {}),
            'last_sync': datetime.now().isoformat()
        }
        with open('.session_data.json', 'w') as f:
            json.dump(state_data, f)
        logger.info("Session state saved successfully.")
    except Exception as e:
        logger.error(f"Failed to save session state: {e}", exc_info=True)
        st.error(f"Failed to save session state: {str(e)}") # Also show user

def load_session_state():
    """Load saved session state data, clearing old data."""
    session_file = '.session_data.json'
    try:
        if os.path.exists(session_file):
            logger.info(f"Loading session state from {session_file}")
            with open(session_file, 'r') as f:
                state_data = json.load(f)
            
            st.session_state.priority_overrides = state_data.get('priority_overrides', {})
            logger.info(f"Loaded {len(st.session_state.priority_overrides)} priority overrides.")
            
            last_sync_str = state_data.get('last_sync')
            if last_sync_str:
                last_sync = datetime.fromisoformat(last_sync_str)
                if datetime.now() - last_sync > timedelta(hours=24):
                    logger.info("Saved session state is older than 24 hours. Clearing.")
                    os.remove(session_file)
                    st.session_state.priority_overrides = {}
            else:
                logger.warning("No last_sync timestamp found in session data. Clearing.")
                os.remove(session_file)
                st.session_state.priority_overrides = {}
        else:
            logger.info("No saved session state file found.")
            st.session_state.priority_overrides = {} # Ensure it's initialized

    except FileNotFoundError:
        logger.info(f"Session state file {session_file} not found (normal on first run).")
        st.session_state.priority_overrides = {} # Ensure initialized
    except json.JSONDecodeError as e:
        logger.error(f"Error decoding JSON from session file {session_file}: {e}. Clearing file.")
        try:
            os.remove(session_file)
        except OSError as remove_err:
            logger.error(f"Failed to remove corrupted session file: {remove_err}")
        st.session_state.priority_overrides = {}
        st.error("Failed to load previous session data (file corrupted). Starting fresh.")
    except Exception as e:
        logger.error(f"Failed to load session state: {e}", exc_info=True)
        st.error(f"Failed to load session state: {str(e)}")
        st.session_state.priority_overrides = {} # Reset on error

def initialize_session_state():
    """Initialize all required session state variables if they don't exist."""
    # Central place to define default state values
    defaults = {
        'authenticated': False,
        'selected_service': None,
        'emails': [],
        'page': 1,
        'total_emails': 0,
        'emails_per_page': 20, # Configurable? Maybe later.
        'all_emails_loaded': False,
        'last_fetched_count': 0,
        'emails_loaded': False,
        'form_key': f"login_{int(time.time())}", # More specific key
        'priority_overrides': {},
        'email_client': None, # Initialize client only when needed
        'action_in_progress': False # Flag to disable buttons during actions
    }
    changed = False
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value
            changed = True
            logger.debug(f"Initialized session state key: {key}")
    if changed:
        logger.info("Session state initialized.")
    
    # Ensure email client is created if authenticated but somehow missing
    if st.session_state.authenticated and st.session_state.email_client is None:
        logger.warning("Authenticated state inconsistency: Re-initializing EmailClient.")
        st.session_state.email_client = EmailClient()

# --- UI Components & Popups --- #

def show_confirmation_dialog(title, message, on_confirm):
    """Show a confirmation dialog using inline containers instead of st.dialog."""
    logger.debug(f"Showing confirmation dialog: {title}")
    # Add unique keys using title hash to prevent conflicts if multiple dialogs were possible
    dialog_key_base = hash(title)
    
    # Create a container for the dialog
    st.markdown(f"### {title}")
    st.write(message)
    col1, col2 = st.columns(2)
    with col1:
        if st.button("Cancel", key=f"confirm_cancel_{dialog_key_base}", use_container_width=True):
            # Do nothing on cancel, just return
            pass
    with col2:
        # Disable confirm button if an action is already in progress
        if st.button("Confirm", key=f"confirm_ok_{dialog_key_base}", type="primary", use_container_width=True, disabled=st.session_state.get('action_in_progress', False)):
            st.session_state.action_in_progress = True # Set flag
            try:
                on_confirm() # Execute the confirmed action
            finally:
                # Ensure flag is reset even if action fails
                st.session_state.action_in_progress = False

def show_error_popup(message):
    """Show an error message using st.popover."""
    logger.warning(f"Showing error popup: {message}") # Log user-facing errors as warnings
    with st.popover("❌ Error"):
        st.error(message)

def show_info_popup(title, message):
    """Show an informational message using st.popover."""
    logger.info(f"Showing info popup: {title} - {message}")
    with st.popover(title):
        st.info(message)

# --- Email Display & Actions --- #

def format_date(date_str):
    """Format date string to a more readable relative format (e.g., '2 hours ago')."""
    if not date_str:
        return "Unknown Date"
    try:
        # Remove potential problematic parenthetical timezone info
        # e.g. "Fri, 11 Apr 2025 19:31:58 +0000 (UTC)" -> "Fri, 11 Apr 2025 19:31:58 +0000"
        clean_date_str = re.sub(r'\([^)]*\)', '', str(date_str)).strip()
        
        # Common email date formats
        formats_to_try = [
            '%a, %d %b %Y %H:%M:%S %z',  # Standard RFC 5322 format
            '%a, %d %b %Y %H:%M:%S %Z',  # Format with timezone name
            '%d %b %Y %H:%M:%S %z',      # Another common variant
            '%Y-%m-%dT%H:%M:%S%z',       # ISO 8601 like format
            '%a, %d %b %Y %H:%M:%S',     # Format without timezone
            '%a, %d %b %Y %H:%M:%S',     # Fallback format with no timezone
        ]
        
        date_obj = None
        for fmt in formats_to_try:
            try:
                date_obj = datetime.strptime(clean_date_str, fmt)
                logger.debug(f"Parsed date '{clean_date_str}' with format '{fmt}'")
                break # Success
            except ValueError:
                continue # Try next format
        
        # Last resort for special cases
        if date_obj is None:
            # Try more parsing techniques
            try:
                # Sometimes there's timezone info in a different format
                # e.g., "Fri, 11 Apr 2025 14:25:49 -0400 (EDT)"
                if 'EDT' in date_str:
                    clean_date_str = re.sub(r'-0400 \(EDT\)', '-0400', clean_date_str)
                    date_obj = datetime.strptime(clean_date_str, '%a, %d %b %Y %H:%M:%S %z')
                elif 'GMT' in date_str:
                    clean_date_str = re.sub(r'\+0000 \(GMT\)', '+0000', clean_date_str)
                    date_obj = datetime.strptime(clean_date_str, '%a, %d %b %Y %H:%M:%S %z')
                elif 'UTC' in date_str:
                    clean_date_str = re.sub(r'\+0000 \(UTC\)', '+0000', clean_date_str)
                    date_obj = datetime.strptime(clean_date_str, '%a, %d %b %Y %H:%M:%S %z')
            except ValueError:
                # If all else fails, try a truly generic approach
                logger.warning(f"Using fallback parsing method for date: {date_str}")
                import dateutil.parser
                try:
                    date_obj = dateutil.parser.parse(clean_date_str)
                except:
                    logger.error(f"Could not parse date string: {date_str} with any known format.")
                    return str(date_str) # Fallback to original string
        
        # Ensure we have a valid date_obj
        if date_obj is None:
            logger.error(f"Could not parse date string: {date_str} with any known format.")
            return str(date_str) # Fallback to original string

        # Ensure we compare timezone-aware with timezone-aware
        if date_obj.tzinfo:
             now = datetime.now(date_obj.tzinfo)
        else:
             # If parsed date is naive, compare with naive now
             now = datetime.now() 

        diff = now - date_obj
        
        # Improved relative time formatting
        if diff.total_seconds() < 0: return "In the future?" # Sanity check
        if diff < timedelta(minutes=1): return "Just now"
        if diff < timedelta(hours=1): 
            minutes = int(diff.total_seconds() / 60)
            return f"{minutes} min ago"
        if diff < timedelta(days=1):
            hours = int(diff.total_seconds() / 3600)
            return f"{hours} hr ago"
        if diff < timedelta(days=7):
            days = diff.days
            return f"{days} day ago" if days == 1 else f"{days} days ago"
        else:
            return date_obj.strftime("%b %d, %Y") # Older dates: just date

    except Exception as e:
        logger.error(f"Unexpected error formatting date '{date_str}': {e}", exc_info=True)
        return str(date_str) # Fallback safely

def get_priority_class(priority):
    """Get CSS class name based on priority level for styling."""
    return f"priority-{priority.lower()}"

def display_email(email, display_idx):
    """Display a single email within an expander, including action buttons."""
    # Use .get() for safer access to potentially missing keys
    priority = email.get('priority', 'Medium')
    priority_class = get_priority_class(priority)
    email_id = email.get('id', f'unknown_{display_idx}') # Use index if ID missing
    subject = email.get('subject', 'No Subject')
    sender = email.get('sender', 'Unknown Sender')
    date_str = email.get('date', None)
    category = email.get('category', 'Other')
    snippet = email.get('snippet', '...')
    
    # Generate unique base key for this email's widgets using display index
    base_key = f"email_{email_id}_{display_idx}"

    # --- Email Header and Metadata --- #
    st.markdown(f"<div class='email-header'>{subject}</div>", unsafe_allow_html=True)
    st.markdown(f"<div class='email-metadata'>From: {sender}</div>", unsafe_allow_html=True)
    st.markdown(f"<div class='email-metadata'>Date: {format_date(date_str)}</div>", unsafe_allow_html=True)
    st.markdown(f"<div class='email-metadata'>Category: {category} • Priority: <span class='{priority_class}'>{priority}</span></div>", unsafe_allow_html=True)
    st.markdown("<hr>", unsafe_allow_html=True)
    # Use st.text or st.markdown for snippet - avoid potential HTML injection if snippet is untrusted
    st.text(snippet)
    
    # --- Priority Override Buttons --- #
    st.markdown("**Change Priority:**")
    prio_cols = st.columns([1, 1, 1, 3]) # Use 4 columns, last one is spacer
    action_disabled = st.session_state.get('action_in_progress', False)
    with prio_cols[0]:
        if st.button("📈 High", key=f"high_{base_key}", use_container_width=True, disabled=action_disabled):
            show_confirmation_dialog(
                f"Confirm Priority: High##{base_key}", # Unique title for dialog instance
                f"Set priority to High for: '{subject}'?",
                lambda eid=email_id: handle_priority_change(eid, "High") # Pass email_id safely
            )
    with prio_cols[1]:
        if st.button("📊 Medium", key=f"medium_{base_key}", use_container_width=True, disabled=action_disabled):
             show_confirmation_dialog(
                f"Confirm Priority: Medium##{base_key}",
                f"Set priority to Medium for: '{subject}'?",
                lambda eid=email_id: handle_priority_change(eid, "Medium")
            )
    with prio_cols[2]:
        if st.button("📉 Low", key=f"low_{base_key}", use_container_width=True, disabled=action_disabled):
            show_confirmation_dialog(
                f"Confirm Priority: Low##{base_key}",
                f"Set priority to Low for: '{subject}'?",
                lambda eid=email_id: handle_priority_change(eid, "Low")
            )
    
    # --- Quick Action Buttons --- #
    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown("**Actions:**")
    action_cols = st.columns(4)
    with action_cols[0]: # Reply
        if st.button("↩️ Reply", key=f"reply_{base_key}", use_container_width=True, disabled=action_disabled):
            st.session_state.action_in_progress = True
            try:
                service = st.session_state.selected_service
                url = get_email_action_url(service, email_id, subject, 'reply')
                if url:
                    open_link(url)
                    logger.info(f"Opened reply link for email {email_id}")
                else:
                    logger.error(f"Could not generate reply URL for {service}, email {email_id}.")
                    st.toast(f"Could not generate reply URL for {service}.", icon="❌")
            except Exception as e:
                logger.error(f"Failed to open reply window for email {email_id}: {e}", exc_info=True)
                st.toast(f"Error opening reply: {str(e)}", icon="❌")
            finally:
                 st.session_state.action_in_progress = False
                 st.rerun() # Rerun to re-enable buttons
    
    with action_cols[1]: # Archive (Placeholder)
        if st.button("📥 Archive", key=f"archive_{base_key}", use_container_width=True, disabled=action_disabled):
            # Placeholder - implement actual archive logic later
            logger.info(f"Archive button clicked for email {email_id} (not implemented)")
            show_info_popup("Coming Soon", "Archive functionality will be added later.")
    
    with action_cols[2]: # Mark Read (Placeholder)
        if st.button("👁️ Mark Read", key=f"read_{base_key}", use_container_width=True, disabled=action_disabled):
            # Placeholder - implement actual mark read logic later
            logger.info(f"Mark Read button clicked for email {email_id} (not implemented)")
            show_info_popup("Coming Soon", "Mark as read functionality will be added later.")
    
    with action_cols[3]: # Open In...
        service = st.session_state.selected_service
        if st.button(f"🔗 Open in {service}", key=f"open_{base_key}", use_container_width=True, disabled=action_disabled):
            st.session_state.action_in_progress = True
            try:
                url = get_email_action_url(service, email_id, subject, 'read')
                if url:
                    open_link(url)
                    logger.info(f"Opened read link for email {email_id} in {service}")
                else:
                    logger.error(f"Could not generate open URL for {service}, email {email_id}.")
                    st.toast(f"Could not generate URL to open in {service}.", icon="❌")
            except Exception as e:
                logger.error(f"Failed to open email {email_id} in {service}: {e}", exc_info=True)
                st.toast(f"Error opening in {service}: {str(e)}", icon="❌")
            finally:
                st.session_state.action_in_progress = False
                st.rerun() # Rerun to re-enable buttons

def get_email_action_url(service, email_id, subject=None, action='read'):
    """Generate the correct URL for Gmail or Outlook based on the action."""
    logger.debug(f"Generating URL for service={service}, action={action}, email_id={email_id}")
    if not email_id or email_id.startswith('unknown_'):
        logger.warning(f"Cannot generate URL for invalid/missing email ID: {email_id}")
        return None
        
    # Clean the email ID just in case
    cleaned_id = str(email_id).strip('<>') 
    
    try:
        if service == "Gmail":
            # Gmail typically uses hex IDs for deep links
            # Add validation or ensure ID format is correct if issues persist
            if action == 'read':
                url = f"https://mail.google.com/mail/u/0/popout/inbox/{cleaned_id}"
            elif action == 'reply':
                url = f"https://mail.google.com/mail/u/0/popout/inbox/r/{cleaned_id}"
            else:
                url = None
        elif service == "Outlook":
            encoded_id = urllib.parse.quote(cleaned_id, safe='') # URL encode the ID
            if action == 'read':
                # Format based on common Outlook Web App URL structure
                url = f"https://outlook.office.com/mail/inbox/id/{encoded_id}"
            elif action == 'reply':
                encoded_subject = urllib.parse.quote(f"Re: {subject}" if subject else "Re:")
                # Use ItemID based on documentation hints, might need adjustment
                url = f"https://outlook.office.com/mail/deeplink/compose?ItemID={encoded_id}&subject={encoded_subject}" 
            else:
                url = None
        else:
            logger.error(f"Invalid service specified for URL generation: {service}")
            url = None
            
        if url:
            logger.info(f"Generated URL: {url}")
        else:
             logger.warning(f"Could not generate URL for action '{action}' on service '{service}'")
        return url
    except Exception as e:
        logger.error(f"Error generating URL: {e}", exc_info=True)
        return None

def open_link(url):
    """Open link in a new tab using JavaScript (works in Streamlit hosted apps)."""
    if not url:
        logger.warning("Attempted to open a null or empty URL.")
        return
    try:
        # Use st.markdown with unsafe_allow_html for script injection
        # Using target="_blank" is standard for new tabs
        # Adding rel="noopener noreferrer" is good practice for security
        st.markdown(f'<a href="{url}" target="_blank" rel="noopener noreferrer" id="openlink" style="display:none;">Open</a><script>document.getElementById("openlink").click();</script>', unsafe_allow_html=True)
        # Alternative JS approach (less reliable across browsers/settings?)
        # js = f'window.open("{url}", "_blank").focus();'
        # st.components.v1.html(f"<script>{js}</script>", height=0)
        logger.info(f"Triggered JavaScript to open URL: {url}")
    except Exception as e:
        logger.error(f"Error trying to open link {url} via JavaScript: {e}", exc_info=True)
        # Fallback for local execution if JS fails?
        # try: webbrowser.open(url)
        # except: pass 
        st.toast("Failed to trigger link opening.", icon="❌")

def update_email_priority(email_id, new_priority):
    """Update the priority of an email and store the override."""
    logger.info(f"Updating priority for email {email_id} to {new_priority}")
    email_found = False
    if 'emails' not in st.session_state or not isinstance(st.session_state.emails, list):
         logger.error("Cannot update priority: Session state 'emails' is not a list.")
         return False
         
    for email in st.session_state.emails:
        # Check if email is a dict and has an 'id' key
        if isinstance(email, dict) and email.get('id') == email_id:
            email['priority'] = new_priority
            email_found = True
            break
    
    if email_found:
        # Store the override in session state
        if 'priority_overrides' not in st.session_state:
             st.session_state.priority_overrides = {}
        st.session_state.priority_overrides[email_id] = new_priority
        logger.info(f"Stored priority override for email {email_id}: {new_priority}")
        return True
    else:
        logger.warning(f"Could not find email with ID {email_id} in current list to update priority.")
        st.warning(f"Could not find email ID {email_id} to update priority.")
        return False

def handle_priority_change(email_id, new_priority):
    """Callback function to handle priority changes, including state saving and rerunning."""
    logger.debug(f"Handling priority change for {email_id} to {new_priority}")
    try:
        success = update_email_priority(email_id, new_priority)
        if success:
            save_session_state() # Save overrides immediately
            st.toast(f"Priority updated to {new_priority}", icon="✅")
            time.sleep(0.5) # Short delay to allow toast to show before rerun
            st.rerun()
        # If not successful, message already shown by update_email_priority
    except Exception as e:
        logger.error(f"Failed to handle priority change for {email_id}: {e}", exc_info=True)
        st.toast(f"Error updating priority: {str(e)}", icon="❌")

def fetch_and_analyze_emails():
    """Fetch emails based on current page, analyze, prioritize, and update session state."""
    if not st.session_state.authenticated:
        logger.warning("Attempted fetch_and_analyze_emails while not authenticated.")
        st.warning("Not authenticated. Please connect to an email service.")
        return
        
    if st.session_state.get('action_in_progress'):
         logger.warning("Fetch attempt blocked: Another action is in progress.")
         return
         
    st.session_state.action_in_progress = True
    logger.info(f"Fetching and analyzing emails for page {st.session_state.page}")
    
    # Ensure email client exists
    if st.session_state.email_client is None:
        logger.error("Email client is None, cannot fetch emails. Re-initializing.")
        st.session_state.email_client = EmailClient()
        # Maybe force a re-auth or show error?
        st.error("Email client connection lost. Please try refreshing or reconnecting.")
        st.session_state.action_in_progress = False
        return
        
    offset = (st.session_state.page - 1) * st.session_state.emails_per_page
    limit = st.session_state.emails_per_page
    fetched_emails = []
    analyzer = None
    scorer = None
    
    try:
        # --- Fetching --- #
        with st.spinner(f"📥 Fetching emails (page {st.session_state.page})..."):
            logger.info(f"Connecting to {st.session_state.selected_service} client to fetch emails...")
            if st.session_state.selected_service == "Gmail":
                fetched_emails = st.session_state.email_client.get_gmail_emails(offset, limit)
            else: # Outlook
                fetched_emails = st.session_state.email_client.get_outlook_emails(offset, limit)
            logger.info(f"Fetched {len(fetched_emails)} emails from {st.session_state.selected_service}.")

        st.session_state.last_fetched_count = len(fetched_emails)
        
        if not fetched_emails:
            if st.session_state.page == 1:
                logger.info("No emails found in inbox.")
                st.toast("No emails found in inbox.", icon="📪")
            else:
                logger.info("No more emails to load.")
                st.toast("No more emails found.", icon="✅")
            st.session_state.all_emails_loaded = True
            # Don't clear existing emails if loading page > 1 resulted in none
            if st.session_state.page == 1: st.session_state.emails = []
            st.session_state.emails_loaded = True # Mark as loaded even if empty
            st.session_state.action_in_progress = False
            st.rerun() # Rerun to show message and potentially hide button
            return

        # --- Analysis & Prioritization --- #
        with st.spinner("🔬 Analyzing and prioritizing emails..."):
            logger.info("Initializing EmailAnalyzer and PriorityScorer...")
            analyzer = EmailAnalyzer()
            scorer = PriorityScorer()
            logger.info(f"Analyzing {len(fetched_emails)} emails...")
            analyzed_emails = analyzer.analyze_emails(fetched_emails)
            logger.info(f"Prioritizing {len(analyzed_emails)} emails...")
            prioritized_emails = scorer.score_emails(analyzed_emails)
            
            logger.info("Applying priority overrides...")
            for email in prioritized_emails:
                if email.get('id') in st.session_state.priority_overrides:
                    original_prio = email.get('priority')
                    override_prio = st.session_state.priority_overrides[email['id']]
                    email['priority'] = override_prio
                    logger.debug(f"Applied override for {email.get('id')}: {original_prio} -> {override_prio}")

        # --- Update Session State --- #
        logger.info("Updating session state with new emails...")
        if st.session_state.page == 1:
            st.session_state.emails = prioritized_emails
        else:
            existing_ids = {email.get('id') for email in st.session_state.emails if email.get('id')}
            new_emails = [email for email in prioritized_emails if email.get('id') not in existing_ids]
            st.session_state.emails.extend(new_emails)
            logger.info(f"Appended {len(new_emails)} new emails to session state.")
        
        if st.session_state.last_fetched_count < limit:
            logger.info("Last fetched count less than limit, assuming all emails loaded.")
            st.session_state.all_emails_loaded = True
        
        st.session_state.total_emails = len(st.session_state.emails)
        st.session_state.emails_loaded = True
        logger.info(f"Fetch/analyze complete. Total emails in state: {st.session_state.total_emails}")
        st.toast("Emails loaded and analyzed!", icon="✨")

    # --- Error Handling for Fetch/Analyze --- #
    except Exception as e:
        logger.error(f"Error during email fetch/analysis: {e}", exc_info=True)
        st.error(f"An error occurred while fetching or analyzing emails: {str(e)}")
        # Optionally display traceback for debugging in UI
        with st.expander("Show Error Details"):
            st.code(traceback.format_exc())
        # Reset state carefully on error?
        st.session_state.last_fetched_count = 0 
        # Don't set all_emails_loaded to True on error
        # Keep existing emails unless it was page 1?
        if st.session_state.page == 1: st.session_state.emails = []

    # --- Finalization --- #
    finally:
        # Ensure the action flag is always reset
        st.session_state.action_in_progress = False
        logger.debug("Action in progress flag reset.")
        # Rerun needed to reflect changes / potentially hide spinner
        st.rerun()

# --- Login & Authentication --- #

EMAIL_REGEX = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"

def create_login_form():
    """Render the login form elements within a Streamlit form."""
    logger.debug("Rendering login form.")
    # Hidden form for potential browser autofill help
    st.markdown("""
        <form id="login-form" style="display:none;">
            <input type="email" name="username" autocomplete="username">
            <input type="password" name="current-password" autocomplete="current-password">
        </form>
    """, unsafe_allow_html=True)

    # Use Streamlit's form for submission handling
    with st.form(key=st.session_state.form_key): # Use dynamic key from session state
        email = st.text_input(
            "Email Address",
            key=f"email_input_{st.session_state.form_key}",
            help="Enter your email address",
            autocomplete="username"
        ).strip()
        
        password = st.text_input(
            "Password / App Password",
            type="password",
            key=f"password_input_{st.session_state.form_key}",
            help="Use App Password if 2FA is enabled",
            autocomplete="current-password"
        ).strip()
        
        submitted = st.form_submit_button("🔒 Connect", disabled=st.session_state.get('action_in_progress', False))
        
        if submitted:
            logger.info(f"Login form submitted for service: {st.session_state.selected_service}")
            st.session_state.action_in_progress = True # Disable buttons during connect
            is_valid, error_message = validate_credentials(email, password)
            if is_valid:
                service = st.session_state.selected_service
                attempt_connection(service, email, password) # This will rerun on success/failure
            else:
                logger.warning(f"Login validation failed: {error_message}")
                st.error(error_message)
                st.session_state.action_in_progress = False # Re-enable form
                # Force form re-render by changing key ONLY on validation error
                st.session_state.form_key = f"login_{int(time.time())}" 
                st.rerun() 

def validate_credentials(email, password):
    """Validate email format and presence of both fields."""
    if not email:
        return False, "❌ Email address is required."
    if not re.match(EMAIL_REGEX, email):
         return False, "❌ Please enter a valid email address format."
    if not password:
        return False, "❌ Password / App Password is required."
    return True, ""

def attempt_connection(service, email, password):
    """Attempt to connect to the selected email service."""
    logger.info(f"Attempting connection to {service} for user {email[:3]}...")
    # Ensure client is initialized before connecting
    if st.session_state.email_client is None:
        st.session_state.email_client = EmailClient()
        logger.info("Initialized EmailClient before connection attempt.")
        
    with st.spinner(f"Connecting to {service}..."):
        connected = False
        try:
            if service == "Gmail":
                connected = st.session_state.email_client.connect_gmail(email, password)
            elif service == "Outlook":
                connected = st.session_state.email_client.connect_outlook(email, password)
            else:
                 logger.error(f"Invalid service selected for connection: {service}")
                 st.error(f"Invalid service selected: {service}")
                 return # Should not happen with radio buttons

            if connected:
                logger.info(f"Successfully connected to {service} for user {email[:3]}...")
                st.session_state.service = service
                st.session_state.user_email = email
                st.session_state.authenticated = True
                st.rerun() # Rerun to reflect state change (connected or error shown)
            else:
                # Specific client errors should ideally be raised and logged by email_client.py
                logger.warning(f"Connection failed for {service} user {email[:3]}... Credentials likely invalid or 2FA/App Password issue.")
                st.error(f"❌ Connection Failed. Check credentials. Use App Password if 2FA is on.")
        
        except Exception as e:
            # Catch broader exceptions during connection attempt
            logger.error(f"Connection error during {service} connect for {email[:3]}...: {e}", exc_info=True)
            st.error(f"❌ Connection Error: {str(e)}")
            with st.expander("Show Error Details"):
                 st.code(traceback.format_exc())
                 
        finally:
            # Always reset flag and rerun after attempt
            st.session_state.action_in_progress = False
            # Regenerate form key to allow re-submission after failed attempt
            st.session_state.form_key = f"login_{int(time.time())}" 
            logger.debug("Connection attempt finished, resetting action flag and form key.")
            st.rerun() # Rerun to reflect state change (connected or error shown)

def handle_disconnect():
    """Show confirmation dialog before disconnecting."""
    logger.debug("Disconnect button clicked, showing confirmation.")
    show_confirmation_dialog(
        "Confirm Disconnect",
        "Are you sure you want to disconnect? This will clear your current session.",
        lambda: perform_disconnect()
    )

def perform_disconnect():
    """Perform the actual disconnection and clear relevant session state."""
    logger.info("Performing disconnection.")
    st.session_state.action_in_progress = True # Prevent other actions during disconnect
    try:
        if st.session_state.email_client:
            st.session_state.email_client.disconnect()
        else:
             logger.warning("Disconnect called but email client was already None.")

        # Clear session state for a clean slate
        keys_to_clear = [
            'authenticated', 'selected_service', 'emails', 'page',
            'total_emails', 'all_emails_loaded', 'last_fetched_count',
            'emails_loaded', 'email_client', 'priority_overrides' # Clear overrides too
        ]
        for key in keys_to_clear:
            if key in st.session_state:
                del st.session_state[key]
        logger.info("Cleared session state keys for disconnect.")
        
        # Clear saved state file on disconnect
        session_file = '.session_data.json'
        if os.path.exists(session_file):
            try:
                os.remove(session_file)
                logger.info("Removed saved session data file.")
            except OSError as e:
                logger.error(f"Failed to remove session data file during disconnect: {e}")
            
        st.toast("Successfully disconnected.", icon="✅")
        time.sleep(1)
        
    except Exception as e:
        logger.error(f"Error during disconnection: {e}", exc_info=True)
        st.toast(f"Error disconnecting: {str(e)}", icon="❌")
    finally:
        st.session_state.action_in_progress = False
        st.rerun()

# --- Main Application Layout --- #

def render_sidebar():
    """Render the sidebar content: connection status, filters, stats, disconnect button."""
    with st.sidebar:
        st.write(f"📨 Connected to {st.session_state.get('selected_service', 'N/A')}")
        
        if st.button("🔓 Disconnect", key="disconnect_button", disabled=st.session_state.get('action_in_progress', False)):
            handle_disconnect()

        st.divider()

        # --- Filters --- #
        st.title("📋 Filters")
        priority_filter = st.selectbox(
            "Priority Level",
            ["All", "High", "Medium", "Low"],
            key="priority_filter",
            disabled=st.session_state.get('action_in_progress', False) # Disable during actions
        )
        
        category_filter = st.selectbox(
            "Category",
            ["All", "Work", "Personal", "Newsletters", "Notifications", "Other"],
            key="category_filter",
            disabled=st.session_state.get('action_in_progress', False)
        )

        st.divider()

        # --- Statistics --- #
        if st.session_state.emails:
            st.title("📊 Statistics")
            try:
                total = len(st.session_state.emails)
                high = sum(1 for e in st.session_state.emails if isinstance(e, dict) and e.get('priority') == 'High')
                medium = sum(1 for e in st.session_state.emails if isinstance(e, dict) and e.get('priority') == 'Medium')
                low = sum(1 for e in st.session_state.emails if isinstance(e, dict) and e.get('priority') == 'Low')
                
                st.metric("Displayed Emails", total)
                st.markdown(f"- **High Priority:** {high}")
                st.markdown(f"- **Medium Priority:** {medium}")
                st.markdown(f"- **Low Priority:** {low}")
            except Exception as e:
                 logger.error(f"Error calculating sidebar statistics: {e}", exc_info=True)
                 st.warning("Could not display statistics.")
            
    return priority_filter, category_filter # Return filter values

def render_email_list(priority_filter, category_filter):
    """Render the main email display area, including headers, refresh, list, and pagination."""
    
    # --- Header and Refresh --- #
    col1, col2 = st.columns([3, 1])
    with col1:
        st.subheader("📬 Prioritized Inbox") # Renamed subheader
    with col2:
        if st.button("🔄 Refresh", key="refresh_button", use_container_width=True, disabled=st.session_state.get('action_in_progress', False)):
            logger.info("Refresh button clicked.")
            st.session_state.emails = [] # Clear current emails
            st.session_state.emails_loaded = False
            st.session_state.page = 1
            st.session_state.all_emails_loaded = False
            # Fetch will be triggered by the rerun and emails_loaded being False
            st.rerun()

    # --- Fetch Emails if needed (initial load or after refresh) --- #
    if not st.session_state.emails_loaded and not st.session_state.get('action_in_progress'):
        logger.debug("Emails not loaded, triggering fetch and analyze.")
        fetch_and_analyze_emails()
        # Fetch function now handles the rerun, so we just return to avoid duplicate render attempts
        return 
        
    # --- Display Email List --- #
    if not st.session_state.emails and st.session_state.emails_loaded:
         # Only show if loading is finished and list is empty
         st.info("📪 Your inbox appears empty or no emails match the current filters.")
    elif st.session_state.emails:
        email_container = st.container()
        with email_container:
            # Filter emails based on sidebar selections
            logger.debug(f"Filtering emails with priority='{priority_filter}' and category='{category_filter}'")
            filtered_emails = [
                email for email in st.session_state.emails
                if isinstance(email, dict) and # Basic type check
                   (priority_filter == "All" or email.get('priority') == priority_filter) and 
                   (category_filter == "All" or email.get('category') == category_filter)
            ]
            logger.debug(f"Found {len(filtered_emails)} emails after filtering.")
            
            if not filtered_emails:
                st.info("📪 No emails match your current filters.")
            else:
                st.write(f"Showing {len(filtered_emails)} emails ({st.session_state.total_emails} loaded total):")
                # Display emails with unique keys based on index in the *filtered* list
                for idx, email in enumerate(filtered_emails):
                    expander_key = f"email_expander_{email.get('id', idx)}_{idx}" # Add filtered index
                    with st.expander(f"{email.get('subject', 'No Subject')} - {email.get('priority', 'Medium')}", expanded=False):
                        display_email(email, idx) 
            
            # --- Pagination Control --- #
            st.divider()
            col_page1, col_page2 = st.columns([1,3])
            with col_page1:
                if not st.session_state.all_emails_loaded:
                    if st.button("📥 Load More Emails", key="load_more", use_container_width=True, disabled=st.session_state.get('action_in_progress', False)):
                        logger.info("'Load More' button clicked.")
                        st.session_state.page += 1
                        # Fetch will be triggered by the rerun and emails_loaded being False (implicitly)
                        # Need to ensure emails_loaded is set False or fetch triggered directly?
                        # For now, assume fetch_and_analyze will run again if needed.
                        st.rerun() 
                elif st.session_state.emails: # Only show if emails exist
                    st.success("✅ All emails loaded.")
            with col_page2:
                 # Optional: Display page number or loaded count
                 st.caption(f"Page {st.session_state.page} | {st.session_state.total_emails} emails loaded")
                 
    elif not st.session_state.emails_loaded:
        # If emails are not loaded and action isn't in progress, show spinner maybe?
        st.spinner("Loading emails...") # Might be redundant if fetch handles spinners

def main():
    """Main function to orchestrate the Streamlit app flow."""
    st.title("📧 AI Email Manager")
    st.write("Connect to Gmail or Outlook to automatically categorize and prioritize your inbox.")
    st.divider()

    # --- Initialization --- #
    # Initialize state first, then load saved data
    initialize_session_state()
    load_session_state() 

    # --- Application Flow --- #
    if not st.session_state.authenticated:
        # --- Login Screen --- #
        st.subheader("🔐 Connect Your Email Account")
        
        with st.popover("ℹ️ Connection Info & Help"):
            st.info("""
                **Notes:**
                - Credentials are used only for this session and **not stored** permanently.
                - Use an **App Password** if 2FA is enabled on Gmail/Outlook.
                - [How to get a Gmail App Password](https://support.google.com/accounts/answer/185833)
                - [Outlook App Passwords info](https://support.microsoft.com/en-us/account-billing/using-app-passwords-with-apps-that-don-t-support-two-step-verification-5896ed9b-4263-e681-128a-a6f2979a7944) (Less common now, might just need normal password)
            """)
        
        # Radio button selection for service
        selected_service_on_login = st.radio(
            "Select Email Service:", 
            ["Gmail", "Outlook"], 
            key="service_select_login", # Use a different key than state
            horizontal=True
        )
        # Update session state based on radio button immediately
        st.session_state.selected_service = selected_service_on_login
        
        create_login_form() # Renders the st.form
        
    else:
        # --- Main Application Screen (Authenticated) --- #
        priority_filter, category_filter = render_sidebar()
        render_email_list(priority_filter, category_filter)

# --- Global Exception Handler --- #
if __name__ == "__main__":
    logger.info("===== Application Starting ====")
    try:
        main()
    except Exception as e:
        # Log the full error and show a user-friendly message + details
        logger.critical(f"An unhandled exception occurred in main: {e}", exc_info=True)
        st.error("🚨 An Unexpected Application Error Occurred!")
        st.error(f"Error details: {str(e)}")
        with st.expander("Show Detailed Error Traceback"):
            st.code(traceback.format_exc())
            
        # Offer a way to restart / clear state
        st.warning("You might need to restart the application.")
        if st.button("🔄 Attempt to Restart Application"):
            logger.warning("Restart button clicked after unhandled exception.")
            keys_to_clear = list(st.session_state.keys())
            for key in keys_to_clear:
                del st.session_state[key]
            session_file = '.session_data.json'
            if os.path.exists(session_file): 
                try: os.remove(session_file)
                except OSError as rm_err: logger.error(f"Failed to clear session file on restart: {rm_err}")
            logger.info("Cleared session state and attempting restart.")
            st.rerun()
    logger.info("===== Application Terminated ====") # May not always log if killed abruptly