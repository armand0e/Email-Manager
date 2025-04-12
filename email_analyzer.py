import nltk
import spacy
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize, sent_tokenize
from nltk.stem import WordNetLemmatizer
import re
import logging # Import logging
from collections import Counter
from datetime import datetime
import os

# Configure logging for this module
logger = logging.getLogger(__name__)

# --- NLTK Resource Check --- #
def download_nltk_resource(resource_id, resource_subdir, download_dir=None):
    """Checks if an NLTK resource exists, downloads if not."""
    try:
        # Check if the resource is available without specifying path first
        nltk.data.find(f'{resource_subdir}/{resource_id}')
        logger.info(f"NLTK resource '{resource_id}' ({resource_subdir}) already available.")
    except LookupError:
        logger.warning(f"NLTK resource '{resource_id}' ({resource_subdir}) not found. Attempting download...")
        try:
            # Create directory if it doesn't exist
            if download_dir is None:
                # Use default directory
                download_dir = os.path.join(os.path.expanduser("~"), "nltk_data")
            
            if not os.path.exists(download_dir):
                os.makedirs(download_dir)
                logger.info(f"Created NLTK data directory: {download_dir}")
            
            # Try to download with quiet=False to see potential error messages
            nltk.download(resource_id, download_dir=download_dir, quiet=False)
            
            # Verify download
            nltk.data.find(f'{resource_subdir}/{resource_id}')
            logger.info(f"NLTK resource '{resource_id}' downloaded successfully.")
        except Exception as e:
            logger.error(f"Failed to download NLTK resource '{resource_id}': {e}", exc_info=True)
            # Don't raise exception to allow the app to continue with degraded functionality
            logger.warning(f"The application will continue but {resource_id} functionality will be limited.")

# Ensure necessary NLTK data is downloaded on import or first use
try:
    # Create a custom download directory in the project folder
    project_nltk_data = os.path.join(os.path.dirname(os.path.abspath(__file__)), "nltk_data")
    
    # Make sure the directory exists
    if not os.path.exists(project_nltk_data):
        os.makedirs(project_nltk_data)
    
    # Add our path to NLTK's search path FIRST before attempting to download or find
    if project_nltk_data not in nltk.data.path:
        nltk.data.path.insert(0, project_nltk_data)
    
    # Simple direct download of resources without verification
    nltk.download('punkt', download_dir=project_nltk_data, quiet=False)
    nltk.download('stopwords', download_dir=project_nltk_data, quiet=False)
    nltk.download('wordnet', download_dir=project_nltk_data, quiet=False)
    
    logger.info(f"NLTK resources downloaded to custom path: {project_nltk_data}")
    logger.info(f"Current NLTK data path: {nltk.data.path}")
except Exception as e:
    # Log the error but allow the application to continue with degraded functionality
    logger.warning(f"Failed to ensure all NLTK resources: {e}")
    logger.info("The application will continue with limited NLP functionality.")

# --- End NLTK Resource Check --- #

# Load spaCy model (ensure it's downloaded: python -m spacy download en_core_web_sm)
try:
    nlp = spacy.load("en_core_web_sm")
    logger.info("spaCy model 'en_core_web_sm' loaded successfully.")
except OSError:
    logger.warning("spaCy model 'en_core_web_sm' not found. Will attempt to download it.")
    try:
        # Try to auto-download the model
        import subprocess
        import sys
        
        # Run the download command
        logger.info("Attempting to download spaCy model...")
        subprocess.check_call([sys.executable, "-m", "spacy", "download", "en_core_web_sm"])
        
        # Try loading again
        nlp = spacy.load("en_core_web_sm")
        logger.info("spaCy model downloaded and loaded successfully.")
    except Exception as e:
        logger.error(f"Failed to download spaCy model: {e}", exc_info=True)
        logger.warning("NER features will be disabled. App will continue with limited functionality.")
        nlp = None

class EmailAnalyzer:
    def __init__(self):
        self.lemmatizer = WordNetLemmatizer()
        try:
            self.stop_words = set(stopwords.words('english'))
            logger.info("NLTK stopwords loaded.")
        except LookupError:
            logger.error("Failed to load NLTK stopwords even after download attempt. Preprocessing quality may be reduced.")
            self.stop_words = set() # Use empty set as fallback
        
        # Define more comprehensive category patterns with improved scoring
        self.category_patterns = {
            'Work': {
                'keywords': ['meeting', 'project', 'deadline', 'report', 'team', 'work', 'task', 
                           'presentation', 'review', 'update', 'status', 'call', 'conference',
                           'agenda', 'minutes', 'action items', 'deliverable', 'milestone',
                           'business', 'office', 'workplace', 'colleague', 'manager', 'boss',
                           'department', 'company', 'organization', 'corporate'],
                'patterns': [
                    (r'status update', 4),
                    (r'project update', 4),
                    (r'team meeting', 4),
                    (r'weekly report', 4),
                    (r'monthly review', 4),
                    (r'business proposal', 4),
                    (r'work related', 3),
                    (r'company policy', 3)
                ]
            },
            'Personal': {
                'keywords': ['family', 'friend', 'personal', 'invitation', 'party', 'dinner',
                           'birthday', 'holiday', 'weekend', 'vacation', 'travel', 'trip',
                           'celebration', 'gathering', 'get together', 'social', 'personal',
                           'private', 'family', 'friends', 'social event', 'personal life'],
                'patterns': [
                    (r'invitation', 4),
                    (r'birthday party', 4),
                    (r'family gathering', 4),
                    (r'weekend plans', 3),
                    (r'personal matter', 3),
                    (r'social event', 3)
                ]
            },
            'Newsletters': {
                'keywords': ['newsletter', 'subscribe', 'unsubscribe', 'digest', 'update',
                           'weekly roundup', 'monthly digest', 'news roundup', 'industry news',
                           'market update', 'trends', 'insights', 'subscription', 'news',
                           'updates', 'alerts', 'notifications', 'marketing', 'promotion'],
                'patterns': [
                    (r'newsletter', 4),
                    (r'weekly digest', 4),
                    (r'monthly update', 4),
                    (r'industry news', 4),
                    (r'subscription', 3),
                    (r'marketing email', 3)
                ]
            },
            'Notifications': {
                'keywords': ['notification', 'alert', 'reminder', 'system', 'update',
                           'password', 'security', 'login', 'account', 'verification',
                           'confirmation', 'receipt', 'order', 'payment', 'system',
                           'automated', 'alert', 'warning', 'notice', 'reminder'],
                'patterns': [
                    (r'system notification', 4),
                    (r'security alert', 4),
                    (r'password reset', 4),
                    (r'order confirmation', 4),
                    (r'automated message', 3),
                    (r'system update', 3)
                ]
            },
            'Other': {
                'keywords': [],
                'patterns': []
            }
        }

        # Define urgency indicators with improved weights
        self.urgency_indicators = {
            'urgent': 4,
            'asap': 4,
            'immediate': 4,
            'important': 3,
            'priority': 3,
            'deadline': 3,
            'reminder': 2,
            'attention': 2,
            'critical': 4,
            'emergency': 4
        }

    def _preprocess_text(self, text):
        """Clean, tokenize, and lemmatize text. Returns original cleaned text and processed tokens."""
        if not text:
            return "", [] # Handle empty text
        
        # Basic cleaning
        cleaned_text = str(text).lower() # Ensure string and lowercase
        cleaned_text = re.sub(r'\S+@\S+', ' [EMAIL] ', cleaned_text)  # Replace email addresses
        cleaned_text = re.sub(r'http\S+|www\S+|https\S+', ' [URL] ', cleaned_text) # Replace URLs
        cleaned_text = re.sub(r'[^\w\s\.\',!?-]', '', cleaned_text) # Keep basic punctuation for context
        cleaned_text = re.sub(r'\d+', ' [NUMBER] ', cleaned_text) # Replace numbers
        cleaned_text = re.sub(r'\s+', ' ', cleaned_text).strip() # Normalize whitespace

        tokens = []
        try:
            # Tokenize using NLTK (relies on punkt)
            tokens = word_tokenize(cleaned_text)
        except LookupError:
            logger.error("NLTK 'punkt' resource not available for tokenization. Analysis limited.")
            # Fallback: simple split (less accurate)
            tokens = cleaned_text.split()
        except Exception as e:
             logger.error(f"Error during word tokenization: {e}", exc_info=True)
             tokens = cleaned_text.split() # Fallback

        # Lemmatize and remove stop words (if available)
        processed_tokens = []
        for word in tokens:
             # Check against stopwords (case-insensitive due to lowercasing earlier)
             if word not in self.stop_words and len(word) > 2:
                 try:
                      lemma = self.lemmatizer.lemmatize(word)
                      processed_tokens.append(lemma)
                 except Exception as lem_err:
                      logger.warning(f"Error lemmatizing word '{word}': {lem_err}")
                      processed_tokens.append(word) # Use original word on error
        
        logger.debug(f"Preprocessing finished. Original length: {len(text)}, Tokens: {len(tokens)}, Processed: {len(processed_tokens)}")
        return cleaned_text, processed_tokens

    def _extract_entities(self, text):
        """Extract named entities using spaCy."""
        if not nlp:
            logger.warning("spaCy model not loaded. Skipping Named Entity Recognition.")
            return [] 
        if not text:
             return [] 
        try:
            doc = nlp(text)
            entities = [(ent.text, ent.label_) for ent in doc.ents]
            logger.debug(f"Extracted entities: {entities}")
            return entities
        except Exception as e:
             logger.error(f"Error during Named Entity Recognition: {e}", exc_info=True)
             return []

    def _detect_category(self, text_tokens):
        """Detect email category based on keywords."""
        if not text_tokens:
            return "Other" # Default category for empty content
            
        # Simple keyword-based categorization (can be significantly expanded)
        # Consider moving these to a config file later
        categories = {
            "Work": {"meeting", "project", "report", "deadline", "client", "colleague", "presentation", "invoice", "work", "office", "schedule", "agenda"},
            "Personal": {"family", "friend", "birthday", "party", "dinner", "weekend", "personal", "trip", "home", "hello", "hi"},
            "Newsletters": {"unsubscribe", "newsletter", "update", "promotion", "sale", "weekly", "daily", "digest", "subscription", "offer"},
            "Notifications": {"alert", "notification", "verify", "confirm", "password", "account", "security", "login", "system", "reset", "message", "comment", "reply"}
        }
        
        tokens_set = set(text_tokens)
        detected_category = "Other" # Default
        max_overlap = 0

        # Find category with the most keyword overlaps
        for category, keywords in categories.items():
            overlap = len(tokens_set.intersection(keywords))
            if overlap > max_overlap:
                max_overlap = overlap
                detected_category = category
        
        logger.debug(f"Detected category: {detected_category} (Overlap: {max_overlap})")
        return detected_category

    def analyze_emails(self, emails):
        """Processes a list of email dictionaries, adding analysis fields."""
        if not isinstance(emails, list):
             logger.error("analyze_emails expected a list, got {type(emails)}.")
             return []
             
        analyzed_emails = []
        logger.info(f"Starting analysis for batch of {len(emails)} emails.")
        
        for i, email in enumerate(emails):
            if not isinstance(email, dict):
                logger.warning(f"Skipping item {i} in email list as it is not a dictionary.")
                continue
                
            email_id = email.get('id', f'no_id_{i}')
            subject = email.get('subject', '')
            body = email.get('body', '')
            logger.debug(f"Analyzing email ID: {email_id} - Subject: {subject[:50]}...")
            
            # Ensure content is string
            full_text = f"{str(subject)} {str(body)}"
            
            analysis_data = {
                 'processed_tokens': [],
                 'entities': [],
                 'category': "Other" # Default category
            }
            
            try:
                # Perform analysis steps
                cleaned_text, tokens = self._preprocess_text(full_text)
                analysis_data['processed_tokens'] = tokens
                # Use cleaned_text for NER as it retains more context than just tokens
                analysis_data['entities'] = self._extract_entities(cleaned_text) 
                analysis_data['category'] = self._detect_category(tokens)
            except Exception as e:
                logger.error(f"Error analyzing email ID {email_id}: {e}", exc_info=True)
                # Keep default analysis data on error

            # Create a new dictionary with original email data + analysis results
            analyzed_email = email.copy()
            analyzed_email.update(analysis_data)
            analyzed_emails.append(analyzed_email)
            
        logger.info(f"Finished analysis batch. Processed {len(analyzed_emails)} emails.")
        return analyzed_emails

    def _determine_category(self, text, tokens):
        """Determine the category of an email using both keywords and patterns"""
        category_scores = {category: 0 for category in self.category_patterns}
        
        # Score based on keywords
        for category, patterns in self.category_patterns.items():
            # Check keywords
            for keyword in patterns['keywords']:
                if keyword in text:
                    category_scores[category] += 2  # Increased weight for keyword matches
            
            # Check patterns
            for pattern, weight in patterns['patterns']:
                if re.search(pattern, text):
                    category_scores[category] += weight
        
        # Get the category with highest score
        max_category = max(category_scores.items(), key=lambda x: x[1])
        
        # If no category scored above threshold, default to 'Other'
        if max_category[1] < 3:  # Increased threshold
            return 'Other', 0
        
        return max_category[0], max_category[1]

    def _calculate_urgency_score(self, tokens):
        """Calculate urgency score based on urgency indicators"""
        score = 0
        for token in tokens:
            if token in self.urgency_indicators:
                score += self.urgency_indicators[token]
        
        # Normalize score to 0-10 range
        return min(score, 10) 